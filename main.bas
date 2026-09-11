Option Explicit

Private Const CONSOLE_NAME As String = "管理コンソール"
Private Const CSV_CHARSET As String = "shift_jis"
Private Const DATE_ORDER As String = "MDY"
Private Const BATCH_ROWS As Long = 2000
Private Const CHART_PREFIX As String = "PMV_"
Private Const DEFAULT_HEIGHT_PX As Double = 200
Private Const LAST_DATA_NAME As String = "_PM_LastData"
Private Const LAST_GRAPH_NAME As String = "_PM_LastGraph"
Private mBusy As Boolean

Private Type CounterSpec
    Header As String
    ColumnIndex As Long
    GroupIndex As Long
    LineColor As Long
End Type

Private Type ExcelState
    Saved As Boolean
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

' CSV import button entry point.
Public Function VisualizePerfmonCsv() As Boolean
    Dim selected As Variant, wb As Workbook, ds As Worksheet, gs As Worksheet
    Dim state As ExcelState, started As Double, ended As Double
    Dim suffix As String, dataName As String, graphName As String
    Dim charts As Long, errorText As String, committed As Boolean

    If mBusy Then Exit Function
    On Error GoTo Failed

    selected = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv", _
        Title:="Select a perfmon CSV log", MultiSelect:=False)

    If VarType(selected) = vbBoolean Then Exit Function

    Set wb = ThisWorkbook
    CheckWorkbook wb
    BeginWork state

    Set ds = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ImportCsvData CStr(selected), ds, started, ended

    suffix = Format$(CDate(started), "yyyymmddhhnn") & "-" & _
             Format$(CDate(ended), "yyyymmddhhnn")
    dataName = "data_" & suffix
    graphName = "graph_" & suffix

    Set gs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    charts = RebuildGraphs(ds, gs)

    DeleteSheetIfExists wb, graphName
    DeleteSheetIfExists wb, dataName

    ds.Name = dataName
    gs.Name = graphName
    committed = True

    RememberLastPair ds, gs
    VisualizePerfmonCsv = True
    GoTo CleanUp

Failed:
    errorText = Err.Description

CleanUp:
    On Error Resume Next

    If Not committed Then
        If Not gs Is Nothing Then gs.Delete
        If Not ds Is Nothing Then ds.Delete
    End If

    On Error GoTo 0
    EndWork state

    If Len(errorText) > 0 Then
        MsgBox "Import failed: " & errorText, vbExclamation, "Perfmon"
    ElseIf VisualizePerfmonCsv Then
        MsgBox "Completed: " & CStr(charts) & " charts." & vbCrLf & _
               ds.Name & vbCrLf & gs.Name, vbInformation, "Perfmon"
    End If
End Function

' Refresh button entry point. Never opens or reads a CSV file.
Public Function RefreshPerfmonGraphs() As Boolean
    Dim ds As Worksheet, gs As Worksheet, state As ExcelState
    Dim charts As Long, errorText As String

    If mBusy Then Exit Function
    On Error GoTo Failed

    CheckWorkbook ThisWorkbook
    GetLastPair ds, gs
    BeginWork state

    charts = RebuildGraphs(ds, gs)
    RememberLastPair ds, gs

    RefreshPerfmonGraphs = True
    GoTo CleanUp

Failed:
    errorText = Err.Description

CleanUp:
    EndWork state

    If Len(errorText) > 0 Then
        MsgBox "Refresh failed: " & errorText, vbExclamation, "Perfmon"
    ElseIf RefreshPerfmonGraphs Then
        MsgBox "Updated: " & CStr(charts) & " charts." & vbCrLf & gs.Name, _
               vbInformation, "Perfmon"
    End If
End Function

' Shared by CSV import and refresh. Validate before deleting charts.
Private Function RebuildGraphs(ByVal ds As Worksheet, ByVal gs As Worksheet) As Long
    Dim console As Worksheet, specs() As CounterSpec, labels() As String
    Dim itemCount As Long, groupCount As Long, firstRow As Long, lastRow As Long
    Dim widthPx As Double, heightPx As Double
    Dim hasStart As Boolean, hasEnd As Boolean, endExclusive As Boolean
    Dim fromDate As Double, toDate As Double
    Dim width As Double, height As Double, gap As Double

    Set console = GetConsole()

    If gs.ProtectContents Or gs.ProtectDrawingObjects Then
        Fail "The graph sheet is protected."
    End If

    ReadSettings console, specs, itemCount, labels, groupCount
    ReadSize console, widthPx, heightPx

    ReadBoundary console.Range("H24"), False, hasStart, fromDate, endExclusive
    ReadBoundary console.Range("H25"), True, hasEnd, toDate, endExclusive

    If hasStart And hasEnd Then
        If endExclusive Then
            If fromDate >= toDate Then Fail "H24 is later than H25."
        Else
            If fromDate > toDate Then Fail "H24 is later than H25."
        End If
    End If

    ResolveColumns ds, specs, itemCount

    FindPlotRows ds, hasStart, fromDate, hasEnd, toDate, _
                 endExclusive, firstRow, lastRow

    ThisWorkbook.Activate
    gs.Activate

    ActiveWindow.View = xlNormalView
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ChartSize ActiveWindow, widthPx, heightPx, width, height, gap

    ' Delete ALL embedded charts, including charts added manually.
    Do While gs.ChartObjects.Count > 0
        gs.ChartObjects(1).Delete
    Loop

    BuildCharts ds, gs, specs, itemCount, labels, groupCount, _
                firstRow, lastRow, widthPx, heightPx, ActiveWindow

    RebuildGraphs = groupCount
End Function

Private Sub ResolveColumns(ByVal ds As Worksheet, _
                           ByRef specs() As CounterSpec, _
                           ByVal count As Long)
    Dim names As Object, lastColumn As Long
    Dim c As Long, i As Long, header As String

    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbBinaryCompare

    lastColumn = ds.Cells(1, ds.Columns.Count).End(xlToLeft).Column
    If lastColumn < 2 Then Fail "The data sheet has no counter columns."

    For c = 1 To lastColumn
        If IsError(ds.Cells(1, c).Value2) Then Fail "Invalid data header."

        header = CStr(ds.Cells(1, c).Value2)

        If Len(header) = 0 Then Fail "An empty data header was found."
        If names.Exists(header) Then Fail "Duplicate data header: " & header

        names.Add header, c
    Next c

    For i = 1 To count
        If Not names.Exists(specs(i).Header) Then
            Fail "Column not found: " & specs(i).Header
        End If

        specs(i).ColumnIndex = CLng(names(specs(i).Header))

        If specs(i).ColumnIndex = 1 Then
            Fail "The timestamp cannot be a counter."
        End If
    Next i
End Sub

Private Sub FindPlotRows(ByVal ds As Worksheet, _
                         ByVal hasStart As Boolean, _
                         ByVal fromDate As Double, _
                         ByVal hasEnd As Boolean, _
                         ByVal toDate As Double, _
                         ByVal endExclusive As Boolean, _
                         ByRef firstRow As Long, _
                         ByRef lastRow As Long)
    Dim endRow As Long, dates As Variant, i As Long
    Dim stamp As Double, previous As Double, inPeriod As Boolean

    endRow = ds.Cells(ds.Rows.Count, 1).End(xlUp).Row
    If endRow < 2 Then Fail "The data sheet has no samples."

    ' Reading A1 too guarantees a two-dimensional array even for one sample.
    dates = ds.Range(ds.Cells(1, 1), ds.Cells(endRow, 1)).Value2

    firstRow = 0
    lastRow = 0

    For i = 2 To endRow
        If IsError(dates(i, 1)) Or IsEmpty(dates(i, 1)) Then
            Fail "Invalid timestamp at row " & CStr(i)
        End If

        If Not IsNumeric(dates(i, 1)) Then
            Fail "Timestamp is not an Excel date at row " & CStr(i)
        End If

        stamp = CDbl(dates(i, 1))

        If stamp < CDbl(DateSerial(1900, 3, 1)) Or stamp >= 2958466# Then
            Fail "Timestamp out of range at row " & CStr(i)
        End If

        If i > 2 And stamp < previous Then
            Fail "Timestamps are out of order at row " & CStr(i)
        End If

        previous = stamp
        inPeriod = True

        If hasStart Then
            If stamp < fromDate Then inPeriod = False
        End If

        If hasEnd Then
            If endExclusive Then
                If stamp >= toDate Then inPeriod = False
            Else
                If stamp > toDate Then inPeriod = False
            End If
        End If

        If inPeriod Then
            If firstRow = 0 Then firstRow = i
            lastRow = i
        End If
    Next i

    If firstRow = 0 Then Fail "No samples fall within H24:H25."
End Sub

Private Sub ImportCsvData(ByVal path As String, _
                          ByVal ds As Worksheet, _
                          ByRef started As Double, _
                          ByRef ended As Double)
    Dim body As String, p As Long
    Dim headers As Variant, fields As Variant, buf() As Variant
    Dim numRx As Object, dateRx As Object
    Dim nCols As Long, total As Long, n As Long, c As Long
    Dim stamp As Double, number As Double

    If DATE_ORDER <> "MDY" And DATE_ORDER <> "DMY" Then
        Fail "Invalid DATE_ORDER."
    End If

    Application.StatusBar = "Reading CSV..."
    body = ReadTextFile(path)

    If Len(body) = 0 Then Fail "The CSV is empty."

    p = 1
    headers = ReadCsvRecord(body, p)
    nCols = UBound(headers) + 1

    If nCols < 2 Then Fail "Expected a timestamp and at least one counter."

    Set numRx = CreateObject("VBScript.RegExp")
    numRx.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([Ee][+-]?[0-9]+)?$"

    Set dateRx = CreateObject("VBScript.RegExp")
    dateRx.IgnoreCase = True
    dateRx.Pattern = "^([0-9]{1,4})[/-]([0-9]{1,2})[/-]([0-9]{1,4})[ T]+([0-9]{1,2}):([0-9]{2}):([0-9]{2})(\.[0-9]+)?[ ]*(AM|PM)?$"

    ds.Rows(1).NumberFormat = "@"

    ReDim buf(1 To 1, 1 To nCols)

    For c = 1 To nCols
        buf(1, c) = headers(c - 1)
    Next c

    ds.Cells(1, 1).Resize(1, nCols).Value2 = buf

    ReDim buf(1 To BATCH_ROWS, 1 To nCols)
    total = 1

    Do While p <= Len(body)
        fields = ReadCsvRecord(body, p)

        If UBound(fields) = 0 Then
            If Len(Trim$(CStr(fields(0)))) = 0 Then GoTo NextRecord
        End If

        If UBound(fields) + 1 <> nCols Then
            Fail "Column count mismatch at row " & CStr(total + 1)
        End If

        If total >= ds.Rows.Count Then
            Fail "The CSV exceeds the Excel row limit."
        End If

        stamp = ParseTimestamp(Trim$(CStr(fields(0))), dateRx, total + 1)

        If total = 1 Then
            started = stamp
        ElseIf stamp < ended Then
            Fail "Timestamps are out of order at row " & CStr(total + 1)
        End If

        ended = stamp
        n = n + 1
        buf(n, 1) = stamp

        For c = 2 To nCols
            If TryNumber(Trim$(CStr(fields(c - 1))), numRx, number) Then
                buf(n, c) = number
            Else
                buf(n, c) = Empty
            End If
        Next c

        total = total + 1

        If n = BATCH_ROWS Then
            ds.Cells(total - n + 1, 1).Resize(n, nCols).Value2 = buf
            n = 0
            ReDim buf(1 To BATCH_ROWS, 1 To nCols)

            Application.StatusBar = _
                "Imported samples: " & Format$(total - 1, "#,##0")
        End If

NextRecord:
    Loop

    If total = 1 Then Fail "No measurement records were found."

    If n > 0 Then
        WriteLastBatch ds, buf, total - n + 1, n, nCols
    End If

    ds.Range(ds.Cells(2, 1), ds.Cells(total, 1)).NumberFormat = _
        "yyyy/mm/dd hh:mm:ss.000"

    ds.Rows(1).Font.Bold = True
    ds.Columns(1).ColumnWidth = 25

    ds.Range(ds.Cells(1, 2), ds.Cells(1, nCols)).EntireColumn.ColumnWidth = 18
    ds.Range(ds.Cells(1, 1), ds.Cells(total, nCols)).AutoFilter
End Sub

Private Sub CheckWorkbook(ByVal wb As Workbook)
    If wb.ProtectStructure Then Fail "The workbook structure is protected."
    If wb.Windows.Count = 0 Then Fail "No workbook window is available."
    If wb.Date1904 Then Fail "Disable the workbook's 1904 date system."
End Sub

Private Sub BeginWork(ByRef state As ExcelState)
    With state
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .DisplayAlerts = Application.DisplayAlerts
        .Calculation = Application.Calculation
        .StatusBar = Application.StatusBar
        .Saved = True
    End With

    mBusy = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndWork(ByRef state As ExcelState)
    On Error Resume Next

    If state.Saved Then
        Application.Calculation = state.Calculation
        Application.EnableEvents = state.EnableEvents
        Application.DisplayAlerts = state.DisplayAlerts
        Application.ScreenUpdating = state.ScreenUpdating
        Application.StatusBar = state.StatusBar
    End If

    mBusy = False
    On Error GoTo 0
End Sub

' Hidden workbook names survive saving/reopening and sheet renaming.
Private Sub RememberLastPair(ByVal ds As Worksheet, ByVal gs As Worksheet)
    ThisWorkbook.Names.Add Name:=LAST_DATA_NAME, _
        RefersTo:="='" & Replace(ds.Name, "'", "''") & "'!$A$1", _
        Visible:=False

    ThisWorkbook.Names.Add Name:=LAST_GRAPH_NAME, _
        RefersTo:="='" & Replace(gs.Name, "'", "''") & "'!$A$1", _
        Visible:=False
End Sub

Private Function StoredSheet(ByVal key As String, _
                             ByRef exists As Boolean) As Worksheet
    Dim nm As Name

    On Error Resume Next
    Set nm = ThisWorkbook.Names(key)
    On Error GoTo 0

    exists = Not (nm Is Nothing)
    If Not exists Then Exit Function

    On Error GoTo Broken

    Set StoredSheet = nm.RefersToRange.Worksheet

    If Not (StoredSheet.Parent Is ThisWorkbook) Then GoTo Broken
    Exit Function

Broken:
    On Error GoTo 0
    Fail "The last output sheet was deleted or its reference is invalid: " & key
End Function

Private Sub GetLastPair(ByRef ds As Worksheet, ByRef gs As Worksheet)
    Dim hasData As Boolean, hasGraph As Boolean
    Dim i As Long, candidate As Worksheet

    Set ds = StoredSheet(LAST_DATA_NAME, hasData)
    Set gs = StoredSheet(LAST_GRAPH_NAME, hasGraph)

    If hasData Or hasGraph Then
        If Not (hasData And hasGraph) Then
            Fail "The last output references are incomplete."
        End If

        If ds Is gs Then
            Fail "The data and graph sheets must be different."
        End If

        Exit Sub
    End If

    ' Migration: the previous version appended its newest graph at the right.
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set candidate = ThisWorkbook.Worksheets(i)

        If Left$(candidate.Name, 6) = "graph_" Then
            Set ds = Nothing

            On Error Resume Next
            Set ds = ThisWorkbook.Worksheets( _
                "data_" & Mid$(candidate.Name, 7))
            On Error GoTo 0

            If Not ds Is Nothing Then
                Set gs = candidate
                Exit Sub
            End If
        End If
    Next i

    Fail "No imported data/graph pair was found. Import a CSV first."
End Sub

Private Function GetConsole() As Worksheet
    On Error Resume Next
    Set GetConsole = ThisWorkbook.Worksheets(CONSOLE_NAME)
    On Error GoTo 0

    If GetConsole Is Nothing Then Fail "Sheet not found: " & CONSOLE_NAME
End Function

Private Sub ReadSettings(ByVal ws As Worksheet, _
                         ByRef specs() As CounterSpec, _
                         ByRef itemCount As Long, _
                         ByRef labels() As String, _
                         ByRef groupCount As Long)
    Dim groups As Object, r As Long
    Dim header As String, key As String
    Dim v As Variant, groupNumber As Double

    Set groups = CreateObject("Scripting.Dictionary")

    ReDim specs(1 To 51)
    ReDim labels(1 To 51)

    For r = 22 To 72
        If IsError(ws.Cells(r, "B").Value2) Then
            Fail "Invalid B" & CStr(r)
        End If

        header = Trim$(CStr(ws.Cells(r, "B").Value2))

        If Len(header) > 0 Then
            v = ws.Cells(r, "C").Value2
            If IsError(v) Then Fail "Invalid C" & CStr(r)

            If Len(Trim$(CStr(v))) = 0 Then
                key = "ROW:" & CStr(r)
            Else
                If Not IsNumeric(v) Then
                    Fail "C" & CStr(r) & " must be a positive integer."
                End If

                groupNumber = CDbl(v)

                If groupNumber < 1 Or groupNumber > 2147483647# Then
                    Fail "Invalid graph number."
                End If

                If groupNumber <> Fix(groupNumber) Then
                    Fail "Graph numbers must be integers."
                End If

                key = "GROUP:" & CStr(CLng(groupNumber))
            End If

            If Not groups.Exists(key) Then
                groupCount = groupCount + 1
                groups.Add key, groupCount

                If Left$(key, 6) = "GROUP:" Then
                    labels(groupCount) = "Graph " & Mid$(key, 7)
                End If
            End If

            itemCount = itemCount + 1

            With specs(itemCount)
                .Header = header
                .GroupIndex = CLng(groups(key))
                .LineColor = CLng(ws.Cells(r, "B").DisplayFormat.Font.Color)
            End With
        End If
    Next r

    If itemCount = 0 Then Fail "Enter counter names in B22:B72."
End Sub

Private Sub ReadSize(ByVal ws As Worksheet, _
                     ByRef widthPx As Double, _
                     ByRef heightPx As Double)
    widthPx = PositivePixels(ws.Range("H22"), 0)
    heightPx = PositivePixels(ws.Range("H23"), DEFAULT_HEIGHT_PX)
End Sub

Private Function PositivePixels(ByVal cell As Range, _
                                ByVal blankDefault As Double) As Double
    Dim v As Variant

    v = cell.Value2

    If IsError(v) Then Fail "Invalid " & cell.Address(False, False)

    If Len(Trim$(CStr(v))) = 0 Then
        PositivePixels = blankDefault
        Exit Function
    End If

    If Not IsNumeric(v) Then
        Fail cell.Address(False, False) & " must contain pixels."
    End If

    If CDbl(v) <= 0 Then
        Fail cell.Address(False, False) & " must be positive."
    End If

    PositivePixels = CDbl(v)
End Function

Private Sub ReadBoundary(ByVal cell As Range, _
                         ByVal isEnd As Boolean, _
                         ByRef supplied As Boolean, _
                         ByRef value As Double, _
                         ByRef exclusive As Boolean)
    Dim v As Variant, text As String
    Dim dateOnly As Boolean, fmt As String, rx As Object

    v = cell.Value2
    exclusive = False

    If IsError(v) Then
        Fail "Invalid date at " & cell.Address(False, False)
    End If

    text = Trim$(CStr(v))
    supplied = (Len(text) > 0)

    If Not supplied Then Exit Sub

    If IsNumeric(v) And VarType(v) <> vbString Then
        value = CDbl(v)

        Set rx = CreateObject("VBScript.RegExp")
        rx.Global = True
        rx.Pattern = """[^""]*""|\\.|\[[^\]]*\]"

        fmt = LCase$(rx.Replace(cell.NumberFormat, ""))

        dateOnly = (value = Fix(value) And _
                    InStr(fmt, "h") = 0 And _
                    InStr(fmt, "s") = 0)
    Else
        If Not IsDate(text) Then
            Fail "Invalid date at " & cell.Address(False, False)
        End If

        value = CDbl(CDate(text))

        dateOnly = (InStr(text, ":") = 0 And _
                    InStr(UCase$(text), "AM") = 0 And _
                    InStr(UCase$(text), "PM") = 0 And _
                    value = Fix(value))
    End If

    If value < CDbl(DateSerial(1900, 3, 1)) Or value >= 2958466# Then
        Fail "Date out of range at " & cell.Address(False, False)
    End If

    If isEnd And dateOnly Then
        value = Fix(value) + 1
        exclusive = True
    End If
End Sub

Private Sub BuildCharts(ByVal ds As Worksheet, _
                        ByVal gs As Worksheet, _
                        ByRef specs() As CounterSpec, _
                        ByVal itemCount As Long, _
                        ByRef labels() As String, _
                        ByVal groupCount As Long, _
                        ByVal firstRow As Long, _
                        ByVal lastRow As Long, _
                        ByVal widthPx As Double, _
                        ByVal heightPx As Double, _
                        ByVal wn As Excel.Window)
    Dim g As Long, i As Long, memberCount As Long
    Dim singleTitle As String
    Dim co As ChartObject, ser As Series, values As Range, x As Range
    Dim maxValue As Double, candidate As Double
    Dim width As Double, height As Double, gap As Double

    ChartSize wn, widthPx, heightPx, width, height, gap

    Set x = ds.Range(ds.Cells(firstRow, 1), ds.Cells(lastRow, 1))

    For g = 1 To groupCount
        Application.StatusBar = _
            "Creating chart " & CStr(g) & "/" & CStr(groupCount)

        Set co = gs.ChartObjects.Add( _
            0, (g - 1) * (height + gap), width, height)

        co.Name = CHART_PREFIX & CStr(g)
        co.Placement = xlFreeFloating

        maxValue = 0
        memberCount = 0

        With co.Chart
            .ChartType = xlLine

            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop

            For i = 1 To itemCount
                If specs(i).GroupIndex = g Then
                    memberCount = memberCount + 1
                    singleTitle = specs(i).Header

                    Set values = ds.Range( _
                        ds.Cells(firstRow, specs(i).ColumnIndex), _
                        ds.Cells(lastRow, specs(i).ColumnIndex))

                    candidate = Application.WorksheetFunction.Max(values)
                    If candidate > maxValue Then maxValue = candidate

                    Set ser = .SeriesCollection.NewSeries

                    ser.Name = "='" & Replace(ds.Name, "'", "''") & "'!" & _
                               ds.Cells(1, specs(i).ColumnIndex).Address
                    ser.XValues = x
                    ser.Values = values
                    ser.MarkerStyle = xlMarkerStyleNone

                    ser.Format.Line.Visible = msoTrue
                    ser.Format.Line.ForeColor.RGB = specs(i).LineColor
                    ser.Format.Line.Weight = 1.5
                    ser.Smooth = False
                End If
            Next i

            .HasTitle = True

            If memberCount = 1 Then
                .ChartTitle.Text = singleTitle
            Else
                .ChartTitle.Text = labels(g)
            End If

            .ChartTitle.Font.Size = 10
            .HasLegend = (memberCount > 1)

            If .HasLegend Then
                .Legend.Position = xlLegendPositionBottom
            End If

            .DisplayBlanksAs = xlNotPlotted
            .PlotVisibleOnly = False

            If maxValue > 0 Then
                maxValue = maxValue * 1.2
            Else
                maxValue = 1
            End If

            With .Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = maxValue
            End With

            With .Axes(xlCategory)
                .CategoryType = xlCategoryScale
                .TickLabels.NumberFormat = "mm/dd hh:mm:ss"
                .TickLabelSpacing = _
                    Application.Min(31999, 1 + (lastRow - firstRow) \ 8)
            End With
        End With
    Next g
End Sub

Private Sub ChartSize(ByVal wn As Excel.Window, _
                      ByVal widthPx As Double, _
                      ByVal heightPx As Double, _
                      ByRef width As Double, _
                      ByRef height As Double, _
                      ByRef gap As Double)
    Dim sx As Double, sy As Double

    ' Differences remove the screen-coordinate offset.
    sx = (CDbl(wn.PointsToScreenPixelsX(720)) - _
          wn.PointsToScreenPixelsX(0)) / 720#

    sy = (CDbl(wn.PointsToScreenPixelsY(720)) - _
          wn.PointsToScreenPixelsY(0)) / 720#

    If sx <= 0 Or sy <= 0 Then
        Fail "Cannot determine the screen scale."
    End If

    If widthPx = 0 Then
        width = wn.UsableWidth / (CDbl(wn.Zoom) / 100#)
    Else
        width = widthPx / sx
    End If

    height = heightPx / sy
    gap = 8# / sy

    If width <= 0 Or height <= 0 Then
        Fail "The Excel window has no usable area."
    End If
End Sub

' Called by ThisWorkbook events. Resize only the displayed graph sheet.
Public Sub ResizePerfmonCharts(ByVal wn As Excel.Window)
    Dim ws As Worksheet, co As ChartObject, console As Worksheet
    Dim widthPx As Double, heightPx As Double
    Dim width As Double, height As Double, gap As Double
    Dim index As Long, found As Boolean

    If mBusy Then Exit Sub
    If wn Is Nothing Then Exit Sub

    On Error GoTo Done

    If wn.WindowState = xlMinimized Then Exit Sub
    If Not TypeOf wn.ActiveSheet Is Worksheet Then Exit Sub

    Set ws = wn.ActiveSheet

    If Not (ws.Parent Is ThisWorkbook) Then Exit Sub
    If Left$(ws.Name, 6) <> "graph_" Then Exit Sub

    For Each co In ws.ChartObjects
        If Left$(co.Name, Len(CHART_PREFIX)) = CHART_PREFIX Then
            found = True
        End If
    Next co

    If Not found Then Exit Sub

    mBusy = True
    Set console = GetConsole()

    ReadSize console, widthPx, heightPx
    ChartSize wn, widthPx, heightPx, width, height, gap

    For Each co In ws.ChartObjects
        If Left$(co.Name, Len(CHART_PREFIX)) = CHART_PREFIX Then
            index = CLng(Mid$(co.Name, Len(CHART_PREFIX) + 1))

            co.Left = 0
            co.Top = (index - 1) * (height + gap)
            co.Width = width
            co.Height = height
        End If
    Next co

Done:
    If Err.Number <> 0 Then
        Debug.Print "Perfmon resize: " & Err.Description
    End If

    mBusy = False
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2100, "PerfmonVisualizer", message
End Sub

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, _
                                ByVal sheetName As String)
    Dim sh As Object

    For Each sh In wb.Sheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            sh.Delete
            Exit Sub
        End If
    Next sh
End Sub

Private Sub WriteLastBatch(ByVal ws As Worksheet, _
                           ByRef source() As Variant, _
                           ByVal firstRow As Long, _
                           ByVal rows As Long, _
                           ByVal cols As Long)
    Dim tail() As Variant, r As Long, c As Long

    ReDim tail(1 To rows, 1 To cols)

    For r = 1 To rows
        For c = 1 To cols
            tail(r, c) = source(r, c)
        Next c
    Next r

    ws.Cells(firstRow, 1).Resize(rows, cols).Value2 = tail
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim stream As Object, bytes As Variant, charset As String
    Dim code As Long, message As String

    On Error GoTo Failed

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path

    charset = CSV_CHARSET

    If stream.Size >= 2 Then
        bytes = stream.Read(3)

        If bytes(0) = &HFF And bytes(1) = &HFE Then
            charset = "unicode"
        End If

        If bytes(0) = &HFE And bytes(1) = &HFF Then
            charset = "unicodeFFFE"
        End If

        If UBound(bytes) >= 2 Then
            If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
                charset = "utf-8"
            End If
        End If
    End If

    stream.Position = 0
    stream.Type = 2
    stream.Charset = charset

    ReadTextFile = stream.ReadText(-1)
    stream.Close

    If Len(ReadTextFile) > 0 Then
        If AscW(Left$(ReadTextFile, 1)) = -257 Then
            ReadTextFile = Mid$(ReadTextFile, 2)
        End If
    End If

    Exit Function

Failed:
    code = Err.Number
    message = Err.Description

    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0

    Err.Raise code, "ReadTextFile", message
End Function

' RFC-style comma parser: quoted commas, escaped quotes, CRLF/LF, embedded newlines.
Private Function ReadCsvRecord(ByRef body As String, _
                               ByRef p As Long) As Variant
    Dim fields() As String, count As Long, capacity As Long
    Dim value As String, ch As String
    Dim quoted As Boolean, closed As Boolean

    capacity = 32
    ReDim fields(0 To capacity - 1)

    Do While p <= Len(body)
        ch = Mid$(body, p, 1)
        p = p + 1

        If quoted Then
            If ch = """" Then
                If Mid$(body, p, 1) = """" Then
                    value = value & """"
                    p = p + 1
                Else
                    quoted = False
                    closed = True
                End If
            Else
                value = value & ch
            End If

        ElseIf ch = "," Or ch = vbCr Or ch = vbLf Then
            PushField fields, count, capacity, value
            value = vbNullString
            closed = False

            If ch <> "," Then
                If ch = vbCr And Mid$(body, p, 1) = vbLf Then p = p + 1

                ReDim Preserve fields(0 To count - 1)
                ReadCsvRecord = fields
                Exit Function
            End If

        ElseIf closed Then
            If ch <> " " And ch <> vbTab Then
                Fail "Unexpected text after a CSV quote."
            End If

        ElseIf ch = """" Then
            If Len(value) <> 0 Then
                Fail "Unexpected quote inside an unquoted CSV field."
            End If

            quoted = True
        Else
            value = value & ch
        End If

        If Len(value) > 32767 Then
            Fail "A CSV field exceeds the Excel cell text limit."
        End If
    Loop

    If quoted Then Fail "Unterminated quoted CSV field."

    PushField fields, count, capacity, value

    ReDim Preserve fields(0 To count - 1)
    ReadCsvRecord = fields
End Function

Private Sub PushField(ByRef fields() As String, _
                      ByRef count As Long, _
                      ByRef capacity As Long, _
                      ByVal value As String)
    If count >= 16384 Then
        Fail "The CSV exceeds the Excel column limit."
    End If

    If count = capacity Then
        capacity = capacity * 2
        ReDim Preserve fields(0 To capacity - 1)
    End If

    fields(count) = value
    count = count + 1
End Sub

Private Function TryNumber(ByVal text As String, _
                           ByVal rx As Object, _
                           ByRef value As Double) As Boolean
    If Len(text) = 0 Then Exit Function
    If Not rx.Test(text) Then Exit Function

    On Error GoTo NotNumeric

    ' Val always uses a period, independently of the Excel/Windows locale.
    value = Val(text)
    TryNumber = True

NotNumeric:
End Function

Private Function ParseTimestamp(ByVal text As String, _
                                ByVal rx As Object, _
                                ByVal row As Long) As Double
    Dim match As Object, yy As Long, mm As Long, dd As Long
    Dim hh As Long, nn As Long, ss As Long
    Dim fraction As Double, datePart As Date, ap As String

    On Error GoTo Invalid

    If Not rx.Test(text) Then GoTo Invalid
    Set match = rx.Execute(text)(0)

    If Len(match.SubMatches(0)) = 4 Then
        yy = CLng(match.SubMatches(0))
        mm = CLng(match.SubMatches(1))
        dd = CLng(match.SubMatches(2))
    Else
        If Len(match.SubMatches(2)) <> 4 Then GoTo Invalid

        yy = CLng(match.SubMatches(2))

        If DATE_ORDER = "MDY" Then
            mm = CLng(match.SubMatches(0))
            dd = CLng(match.SubMatches(1))
        Else
            dd = CLng(match.SubMatches(0))
            mm = CLng(match.SubMatches(1))
        End If
    End If

    If yy < 1900 Or yy > 9999 Or _
       mm < 1 Or mm > 12 Or _
       dd < 1 Or dd > 31 Then GoTo Invalid

    datePart = DateSerial(yy, mm, dd)

    If Year(datePart) <> yy Or _
       Month(datePart) <> mm Or _
       Day(datePart) <> dd Then GoTo Invalid

    ' Modern perfmon logs only; avoids VBA/Excel pre-March-1900 differences.
    If datePart < DateSerial(1900, 3, 1) Then GoTo Invalid

    hh = CLng(match.SubMatches(3))
    nn = CLng(match.SubMatches(4))
    ss = CLng(match.SubMatches(5))
    ap = UCase$(CStr(match.SubMatches(7)))

    If Len(ap) > 0 Then
        If hh < 1 Or hh > 12 Then GoTo Invalid
        hh = hh Mod 12
        If ap = "PM" Then hh = hh + 12
    End If

    If hh > 23 Or nn > 59 Or ss > 59 Then GoTo Invalid

    If Len(match.SubMatches(6)) > 0 Then
        fraction = Val("0" & match.SubMatches(6))
    End If

    ParseTimestamp = CDbl(datePart) + _
        (hh * 3600# + nn * 60# + ss + fraction) / 86400#

    Exit Function

Invalid:
    On Error GoTo 0

    Fail "Unsupported/invalid timestamp at row " & CStr(row) & ": " & text & _
         " (check DATE_ORDER: " & DATE_ORDER & ")."
End Function