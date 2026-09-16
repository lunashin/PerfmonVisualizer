Attribute VB_Name = "PerfmonVisualizer"
Option Explicit

Private Const CONSOLE_NAME As String = "�Ǘ��R���\�[��"
Private Const DAILY_TIME_MARKER As String = "__PMV_DAILY_TIME__"
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
    Description As String
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
    Dim widthPx As Double, heightPx As Double, dailyMode As Boolean
    Dim hasStart As Boolean, hasEnd As Boolean, endExclusive As Boolean
    Dim fromDate As Double, toDate As Double
    Dim width As Double, height As Double, gap As Double
    Dim days() As Double, firsts() As Long, lasts() As Long, times As Variant, dayCount As Long
    Set console = GetConsole()
    If gs.ProtectContents Or gs.ProtectDrawingObjects Then Fail "The graph sheet is protected."
    dailyMode = ReadDailyMode(console)
    ReadSettings console, specs, itemCount, labels, groupCount, dailyMode
    ReadSize console, widthPx, heightPx
    ReadBoundary console.Range("H24"), False, hasStart, fromDate, endExclusive, dailyMode
    ReadBoundary console.Range("H25"), True, hasEnd, toDate, endExclusive, dailyMode
    If hasStart And hasEnd Then
        If endExclusive Then
            If fromDate >= toDate Then Fail "H24 is later than H25."
        Else
            If fromDate > toDate Then Fail "H24 is later than H25."
        End If
    End If
    ResolveColumns ds, specs, itemCount
    FindPlotRows ds, hasStart, fromDate, hasEnd, toDate, endExclusive, firstRow, lastRow
    If dailyMode Then
        CheckDailyTimeColumn gs
        dayCount = GetDaySegments(ds, firstRow, lastRow, days, firsts, lasts, times)
    End If
    ThisWorkbook.Activate
    gs.Activate
    ActiveWindow.View = xlNormalView
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ChartSize ActiveWindow, widthPx, heightPx, width, height, gap
    Do While gs.ChartObjects.Count > 0
        gs.ChartObjects(1).Delete
    Loop
    ClearDailyTimes gs
    If dailyMode Then
        WriteDailyTimes gs, times
        BuildDailyCharts ds, gs, specs, itemCount, firstRow, lastRow, dayCount, _
                         days, firsts, lasts, widthPx, heightPx, ActiveWindow
        RebuildGraphs = itemCount
    Else
        BuildCharts ds, gs, specs, itemCount, labels, groupCount, _
                    firstRow, lastRow, widthPx, heightPx, ActiveWindow
        RebuildGraphs = groupCount
    End If
End Function

Private Sub ResolveColumns(ByVal ds As Worksheet, ByRef specs() As CounterSpec, ByVal count As Long)
    Dim headers() As String, columns() As Long, lastColumn As Long, c As Long, i As Long
    lastColumn = ds.Cells(1, ds.Columns.Count).End(xlToLeft).Column
    If lastColumn < 2 Then Fail "The data sheet has no counter columns."
    ReDim headers(1 To lastColumn)
    ReDim columns(1 To lastColumn)
    For c = 1 To lastColumn
        If IsError(ds.Cells(1, c).Value2) Then Fail "Invalid data header."
        headers(c) = CStr(ds.Cells(1, c).Value2)
        If Len(headers(c)) = 0 Then Fail "An empty data header was found."
        columns(c) = c
    Next c
    SortHeaders headers, columns, 1, lastColumn
    For c = 2 To lastColumn
        If StrComp(headers(c), headers(c - 1), vbBinaryCompare) = 0 Then _
            Fail "Duplicate data header: " & headers(c)
    Next c
    For i = 1 To count
        specs(i).ColumnIndex = FindHeader(headers, columns, specs(i).Header)
        If specs(i).ColumnIndex = 0 Then Fail "Column not found: " & specs(i).Header
        If specs(i).ColumnIndex = 1 Then Fail "The timestamp cannot be a counter."
    Next i
End Sub

Private Sub FindPlotRows(ByVal ds As Worksheet, ByVal hasStart As Boolean, ByVal fromDate As Double, _
    ByVal hasEnd As Boolean, ByVal toDate As Double, ByVal endExclusive As Boolean, _
    ByRef firstRow As Long, ByRef lastRow As Long)
    Dim endRow As Long, dates As Variant, i As Long, stamp As Double, previous As Double
    Dim inPeriod As Boolean
    endRow = ds.Cells(ds.Rows.Count, 1).End(xlUp).Row
    If endRow < 2 Then Fail "The data sheet has no samples."
    ' Reading A1 too guarantees a two-dimensional array even for one sample.
    dates = ds.Range(ds.Cells(1, 1), ds.Cells(endRow, 1)).Value2
    firstRow = 0
    lastRow = 0
    For i = 2 To endRow
        If IsError(dates(i, 1)) Or IsEmpty(dates(i, 1)) Then Fail "Invalid timestamp at row " & CStr(i)
        If Not IsNumeric(dates(i, 1)) Then Fail "Timestamp is not an Excel date at row " & CStr(i)
        stamp = CDbl(dates(i, 1))
        If stamp < CDbl(DateSerial(1900, 3, 1)) Or stamp >= 2958466# Then _
            Fail "Timestamp out of range at row " & CStr(i)
        If i > 2 And stamp < previous Then Fail "Timestamps are out of order at row " & CStr(i)
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

Private Sub ImportCsvData(ByVal path As String, ByVal ds As Worksheet, _
                          ByRef started As Double, ByRef ended As Double)
    Dim body As String, p As Long, headers As Variant, fields As Variant, buf() As Variant
    Dim nCols As Long, total As Long, n As Long, c As Long
    Dim stamp As Double, number As Double
    If DATE_ORDER <> "MDY" And DATE_ORDER <> "DMY" Then Fail "Invalid DATE_ORDER."
    Application.StatusBar = "Reading CSV..."
    body = ReadTextFile(path)
    If Len(body) = 0 Then Fail "The CSV is empty."
    p = 1
    headers = ReadCsvRecord(body, p)
    nCols = UBound(headers) + 1
    If nCols < 2 Then Fail "Expected a timestamp and at least one counter."
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
        If UBound(fields) + 1 <> nCols Then Fail "Column count mismatch at row " & CStr(total + 1)
        If total >= ds.Rows.Count Then Fail "The CSV exceeds the Excel row limit."
        stamp = ParseTimestamp(Trim$(CStr(fields(0))), total + 1)
        If total = 1 Then
            started = stamp
        ElseIf stamp < ended Then
            Fail "Timestamps are out of order at row " & CStr(total + 1)
        End If
        ended = stamp
        n = n + 1
        buf(n, 1) = stamp
        For c = 2 To nCols
            If TryNumber(Trim$(CStr(fields(c - 1))), number) Then
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
            Application.StatusBar = "Imported samples: " & Format$(total - 1, "#,##0")
        End If
NextRecord:
    Loop
    If total = 1 Then Fail "No measurement records were found."
    If n > 0 Then WriteLastBatch ds, buf, total - n + 1, n, nCols
    ds.Range(ds.Cells(2, 1), ds.Cells(total, 1)).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
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
        RefersTo:="='" & Replace(ds.Name, "'", "''") & "'!$A$1", Visible:=False
    ThisWorkbook.Names.Add Name:=LAST_GRAPH_NAME, _
        RefersTo:="='" & Replace(gs.Name, "'", "''") & "'!$A$1", Visible:=False
End Sub

Private Function StoredSheet(ByVal key As String, ByRef exists As Boolean) As Worksheet
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
    Dim hasData As Boolean, hasGraph As Boolean, i As Long, candidate As Worksheet
    Set ds = StoredSheet(LAST_DATA_NAME, hasData)
    Set gs = StoredSheet(LAST_GRAPH_NAME, hasGraph)
    If hasData Or hasGraph Then
        If Not (hasData And hasGraph) Then Fail "The last output references are incomplete."
        If ds Is gs Then Fail "The data and graph sheets must be different."
        Exit Sub
    End If
    ' Migration: the previous version appended its newest graph at the right.
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set candidate = ThisWorkbook.Worksheets(i)
        If Left$(candidate.Name, 6) = "graph_" Then
            Set ds = Nothing
            On Error Resume Next
            Set ds = ThisWorkbook.Worksheets("data_" & Mid$(candidate.Name, 7))
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

Private Sub ReadSettings(ByVal ws As Worksheet, ByRef specs() As CounterSpec, _
    ByRef itemCount As Long, ByRef labels() As String, ByRef groupCount As Long, _
    Optional ByVal dailyMode As Boolean = False)
    Dim groupKeys(1 To 51) As String, r As Long, g As Long, header As String, key As String
    Dim v As Variant, groupNumber As Double
    ReDim specs(1 To 51)
    ReDim labels(1 To 51)
    For r = 22 To 72
        If IsError(ws.Cells(r, "B").Value2) Then Fail "Invalid B" & CStr(r)
        header = Trim$(CStr(ws.Cells(r, "B").Value2))
        If Len(header) > 0 Then
            If IsError(ws.Cells(r, "D").Value2) Then Fail "Invalid D" & CStr(r)
            If dailyMode Then
                v = vbNullString
            Else
                v = ws.Cells(r, "C").Value2
                If IsError(v) Then Fail "Invalid C" & CStr(r)
            End If
            If Len(Trim$(CStr(v))) = 0 Then
                key = "ROW:" & CStr(r)
            Else
                If Not IsNumeric(v) Then Fail "C" & CStr(r) & " must be a positive integer."
                groupNumber = CDbl(v)
                If groupNumber < 1 Or groupNumber > 2147483647# Then Fail "Invalid graph number."
                If groupNumber <> Fix(groupNumber) Then Fail "Graph numbers must be integers."
                key = "GROUP:" & CStr(CLng(groupNumber))
            End If
            For g = 1 To groupCount
                If groupKeys(g) = key Then Exit For
            Next g
            If g > groupCount Then
                groupCount = groupCount + 1
                groupKeys(groupCount) = key
            End If
            itemCount = itemCount + 1
            With specs(itemCount)
                .Header = header
                .Description = Trim$(CStr(ws.Cells(r, "D").Value2))
                .GroupIndex = g
                If Not dailyMode Then
                    .LineColor = CLng(ws.Cells(r, "B").DisplayFormat.Font.Color)
                End If
            End With
            g = specs(itemCount).GroupIndex
            If Len(labels(g)) > 0 Then labels(g) = labels(g) & vbLf
            labels(g) = labels(g) & CounterTitle(specs(itemCount))
        End If
    Next r
    If itemCount = 0 Then Fail "Enter counter names in B22:B72."
End Sub

Private Sub ReadSize(ByVal ws As Worksheet, ByRef widthPx As Double, ByRef heightPx As Double)
    widthPx = PositivePixels(ws.Range("H22"), 0)
    heightPx = PositivePixels(ws.Range("H23"), DEFAULT_HEIGHT_PX)
End Sub

Private Function PositivePixels(ByVal cell As Range, ByVal blankDefault As Double) As Double
    Dim v As Variant
    v = cell.Value2
    If IsError(v) Then Fail "Invalid " & cell.Address(False, False)
    If Len(Trim$(CStr(v))) = 0 Then
        PositivePixels = blankDefault
        Exit Function
    End If
    If Not IsNumeric(v) Then Fail cell.Address(False, False) & " must contain pixels."
    If CDbl(v) <= 0 Then Fail cell.Address(False, False) & " must be positive."
    PositivePixels = CDbl(v)
End Function

Private Sub ReadBoundary(ByVal cell As Range, ByVal isEnd As Boolean, _
    ByRef supplied As Boolean, ByRef value As Double, ByRef exclusive As Boolean, _
    Optional ByVal dateOnlyMode As Boolean = False)
    Dim v As Variant, text As String, dateOnly As Boolean
    v = cell.Value2
    exclusive = False
    If IsError(v) Then Fail "Invalid date at " & cell.Address(False, False)
    text = Trim$(CStr(v))
    supplied = (Len(text) > 0)
    If Not supplied Then Exit Sub
    If IsNumeric(v) And VarType(v) <> vbString Then
        value = CDbl(v)
        dateOnly = (value = Fix(value) And Not FormatHasTime(cell.NumberFormat))
    Else
        If Not IsDate(text) Then Fail "Invalid date at " & cell.Address(False, False)
        value = CDbl(CDate(text))
        dateOnly = (InStr(text, ":") = 0 And InStr(UCase$(text), "AM") = 0 And _
                    InStr(UCase$(text), "PM") = 0 And value = Fix(value))
    End If
    If value < CDbl(DateSerial(1900, 3, 1)) Or value >= 2958466# Then _
        Fail "Date out of range at " & cell.Address(False, False)
    If dateOnlyMode Then
        value = Fix(value)
        dateOnly = True
    End If
    If isEnd And dateOnly Then
        value = Fix(value) + 1
        exclusive = True
    End If
End Sub

Private Function CounterTitle(ByRef spec As CounterSpec) As String
    CounterTitle = spec.Header & "(" & spec.Description & ")"
End Function

' XY charts provide a numeric time axis; 1/24 day is exactly one hour.
Private Sub SetHourlyTimeAxis(ByVal chart As Chart, ByVal firstStamp As Double, _
                              ByVal lastStamp As Double)
    Dim axisMin As Double, axisMax As Double
    axisMin = Int(firstStamp * 24#) / 24#
    axisMax = -Int(-lastStamp * 24#) / 24#
    If axisMax <= axisMin Then axisMax = axisMin + 1# / 24#
    With chart.Axes(xlCategory, xlPrimary)
        .MinimumScale = axisMin
        .MaximumScale = axisMax
        .MajorUnit = 1# / 24#
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormat = "mm/dd hh:mm"
        .TickLabels.Orientation = 45
        .MajorTickMark = xlTickMarkOutside
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionLow
    End With
End Sub

Private Sub BuildCharts(ByVal ds As Worksheet, ByVal gs As Worksheet, _
    ByRef specs() As CounterSpec, ByVal itemCount As Long, ByRef labels() As String, _
    ByVal groupCount As Long, ByVal firstRow As Long, ByVal lastRow As Long, _
    ByVal widthPx As Double, ByVal heightPx As Double, ByVal wn As Excel.Window)
    Dim g As Long, i As Long, memberCount As Long
    Dim co As ChartObject, ser As Series, values As Range, x As Range
    Dim maxValue As Double, candidate As Double, width As Double, height As Double, gap As Double
    ChartSize wn, widthPx, heightPx, width, height, gap
    Set x = ds.Range(ds.Cells(firstRow, 1), ds.Cells(lastRow, 1))
    For g = 1 To groupCount
        Application.StatusBar = "Creating chart " & CStr(g) & "/" & CStr(groupCount)
        Set co = gs.ChartObjects.Add(0, (g - 1) * (height + gap), width, height)
        co.Name = CHART_PREFIX & CStr(g)
        co.Placement = xlFreeFloating
        maxValue = 0
        memberCount = 0
        With co.Chart
            .ChartType = xlXYScatterLinesNoMarkers
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            For i = 1 To itemCount
                If specs(i).GroupIndex = g Then
                    memberCount = memberCount + 1
                    Set values = ds.Range(ds.Cells(firstRow, specs(i).ColumnIndex), _
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
            FinishPlot co.Chart, labels(g), (memberCount > 1), maxValue
            SetHourlyTimeAxis co.Chart, CDbl(ds.Cells(firstRow, 1).Value2), _
                              CDbl(ds.Cells(lastRow, 1).Value2)
        End With
    Next g
End Sub

Private Sub ChartSize(ByVal wn As Excel.Window, ByVal widthPx As Double, _
    ByVal heightPx As Double, ByRef width As Double, ByRef height As Double, ByRef gap As Double)
    Dim sx As Double, sy As Double
    ' Differences remove the screen-coordinate offset.
    sx = (CDbl(wn.PointsToScreenPixelsX(720)) - wn.PointsToScreenPixelsX(0)) / 720#
    sy = (CDbl(wn.PointsToScreenPixelsY(720)) - wn.PointsToScreenPixelsY(0)) / 720#
    If sx <= 0 Or sy <= 0 Then Fail "Cannot determine the screen scale."
    If widthPx = 0 Then
        width = wn.UsableWidth / (CDbl(wn.Zoom) / 100#)
    Else
        width = widthPx / sx
    End If
    height = heightPx / sy
    gap = 8# / sy
    If width <= 0 Or height <= 0 Then Fail "The Excel window has no usable area."
End Sub

' Called by ThisWorkbook events. Resize only the displayed graph sheet.
Public Sub ResizePerfmonCharts(ByVal wn As Excel.Window)
    Dim ws As Worksheet, co As ChartObject, console As Worksheet
    Dim widthPx As Double, heightPx As Double, width As Double, height As Double, gap As Double
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
        If Left$(co.Name, Len(CHART_PREFIX)) = CHART_PREFIX Then found = True
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
    If Err.Number <> 0 Then Debug.Print "Perfmon resize: " & Err.Description
    mBusy = False
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2100, "PerfmonVisualizer", message
End Sub

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            sh.Delete
            Exit Sub
        End If
    Next sh
End Sub

Private Sub WriteLastBatch(ByVal ws As Worksheet, ByRef source() As Variant, _
                           ByVal firstRow As Long, ByVal rows As Long, ByVal cols As Long)
    Dim tail() As Variant, r As Long, c As Long
    ReDim tail(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            tail(r, c) = source(r, c)
        Next c
    Next r
    ws.Cells(firstRow, 1).Resize(rows, cols).Value2 = tail
End Sub

' RFC-style comma parser: quoted commas, escaped quotes, CRLF/LF, embedded newlines.
Private Function ReadCsvRecord(ByRef body As String, ByRef p As Long) As Variant
    Dim fields() As String, count As Long, capacity As Long
    Dim value As String, ch As String, quoted As Boolean, closed As Boolean
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
            If ch <> " " And ch <> vbTab Then Fail "Unexpected text after a CSV quote."
        ElseIf ch = """" Then
            If Len(value) <> 0 Then Fail "Unexpected quote inside an unquoted CSV field."
            quoted = True
        Else
            value = value & ch
        End If
        If Len(value) > 32767 Then Fail "A CSV field exceeds the Excel cell text limit."
    Loop
    If quoted Then Fail "Unterminated quoted CSV field."
    PushField fields, count, capacity, value
    ReDim Preserve fields(0 To count - 1)
    ReadCsvRecord = fields
End Function

Private Sub PushField(ByRef fields() As String, ByRef count As Long, _
                      ByRef capacity As Long, ByVal value As String)
    If count >= 16384 Then Fail "The CSV exceeds the Excel column limit."
    If count = capacity Then
        capacity = capacity * 2
        ReDim Preserve fields(0 To capacity - 1)
    End If
    fields(count) = value
    count = count + 1
End Sub

' Optional wrappers for Windows Form-control buttons.
Public Sub PerfmonImportButton()
    Call VisualizePerfmonCsv
End Sub

Public Sub PerfmonRefreshButton()
    Call RefreshPerfmonGraphs
End Sub

' Sort header names together with their original column numbers.
Private Sub SortHeaders(ByRef headers() As String, ByRef columns() As Long, _
                        ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, tmp As String, tmpCol As Long
    i = first
    j = last
    pivot = headers((first + last) \ 2)
    Do While i <= j
        Do While StrComp(headers(i), pivot, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(headers(j), pivot, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = headers(i): headers(i) = headers(j): headers(j) = tmp
            tmpCol = columns(i): columns(i) = columns(j): columns(j) = tmpCol
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then SortHeaders headers, columns, first, j
    If i < last Then SortHeaders headers, columns, i, last
End Sub

Private Function FindHeader(ByRef headers() As String, ByRef columns() As Long, _
                            ByVal text As String) As Long
    Dim lo As Long, hi As Long, middle As Long, comparison As Long
    lo = LBound(headers)
    hi = UBound(headers)
    Do While lo <= hi
        middle = (lo + hi) \ 2
        comparison = StrComp(headers(middle), text, vbBinaryCompare)
        If comparison = 0 Then
            FindHeader = columns(middle)
            Exit Function
        ElseIf comparison < 0 Then
            lo = middle + 1
        Else
            hi = middle - 1
        End If
    Loop
End Function

Private Function FormatHasTime(ByVal format As String) As Boolean
    Dim i As Long, ch As String, quoted As Boolean, bracketed As Boolean
    For i = 1 To Len(format)
        ch = LCase$(Mid$(format, i, 1))
        If ch = """" Then
            quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "\" Then
                i = i + 1
            ElseIf ch = "[" Then
                bracketed = True
            ElseIf ch = "]" Then
                bracketed = False
            ElseIf Not bracketed Then
                If ch = "h" Or ch = "s" Then
                    FormatHasTime = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function TryNumber(ByVal text As String, ByRef value As Double) As Boolean
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([Ee][+-]?[0-9]+)?$"
    End If
    If Len(text) = 0 Then Exit Function
    If Not rx.Test(text) Then Exit Function
    On Error GoTo NotNumeric
    value = Val(text)
    TryNumber = True
NotNumeric:
End Function

Private Function ParseTimestamp(ByVal text As String, ByVal row As Long) As Double
    Static rx As Object
    Dim match As Object, yy As Long, mm As Long, dd As Long
    Dim hh As Long, nn As Long, ss As Long, fraction As Double, datePart As Date, ap As String
    On Error GoTo Invalid
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.IgnoreCase = True
        rx.Pattern = "^([0-9]{1,4})[/-]([0-9]{1,2})[/-]([0-9]{1,4})[ T]+([0-9]{1,2}):([0-9]{2}):([0-9]{2})(\.[0-9]+)?[ ]*(AM|PM)?$"
    End If
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
    If yy < 1900 Or yy > 9999 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then GoTo Invalid
    datePart = DateSerial(yy, mm, dd)
    If Year(datePart) <> yy Or Month(datePart) <> mm Or Day(datePart) <> dd Then GoTo Invalid
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
    If Len(match.SubMatches(6)) > 0 Then fraction = Val("0" & match.SubMatches(6))
    ParseTimestamp = CDbl(datePart) + (hh * 3600# + nn * 60# + ss + fraction) / 86400#
    Exit Function
Invalid:
    On Error GoTo 0
    Fail "Unsupported/invalid timestamp at row " & CStr(row) & ": " & text & _
         " (check DATE_ORDER: " & DATE_ORDER & ")."
End Function

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
        If bytes(0) = &HFF And bytes(1) = &HFE Then charset = "unicode"
        If bytes(0) = &HFE And bytes(1) = &HFF Then charset = "unicodeFFFE"
        If UBound(bytes) >= 2 Then
            If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then charset = "utf-8"
        End If
    End If
    stream.Position = 0
    stream.Type = 2
    stream.Charset = charset
    ReadTextFile = stream.ReadText(-1)
    stream.Close
    If Len(ReadTextFile) > 0 Then
        If AscW(Left$(ReadTextFile, 1)) = -257 Then ReadTextFile = Mid$(ReadTextFile, 2)
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

Private Function ReadDailyMode(ByVal console As Worksheet) As Boolean
    Dim v As Variant
    v = console.Range("H26").Value2
    If IsError(v) Then Fail "H26 must be 0 or 1."
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then Fail "H26 must be 0 or 1."
    If CDbl(v) <> 0 And CDbl(v) <> 1 Then Fail "H26 must be 0 or 1."
    ReadDailyMode = (CDbl(v) = 1)
End Function

' Split chronologically ordered samples at midnight before drawing.
Private Function GetDaySegments(ByVal ds As Worksheet, ByVal firstRow As Long, _
    ByVal lastRow As Long, ByRef days() As Double, ByRef firsts() As Long, _
    ByRef lasts() As Long, ByRef times As Variant) As Long
    Dim stamps As Variant, count As Long, i As Long, d As Double, stamp As Double
    Dim n As Long
    n = lastRow - firstRow + 1
    If n = 1 Then
        ReDim stamps(1 To 1, 1 To 1)
        stamps(1, 1) = ds.Cells(firstRow, 1).Value2
    Else
        stamps = ds.Range(ds.Cells(firstRow, 1), ds.Cells(lastRow, 1)).Value2
    End If
    ReDim days(1 To 255)
    ReDim firsts(1 To 255)
    ReDim lasts(1 To 255)
    ReDim times(1 To n, 1 To 1)
    For i = 1 To n
        stamp = CDbl(stamps(i, 1))
        d = Fix(stamp)
        If count = 0 Then
            count = 1
            days(count) = d
            firsts(count) = firstRow + i - 1
        ElseIf d <> days(count) Then
            If count = 255 Then Fail "Daily overlay supports at most 255 dates. Shorten H24:H25."
            count = count + 1
            days(count) = d
            firsts(count) = firstRow + i - 1
        End If
        lasts(count) = firstRow + i - 1
        times(i, 1) = stamp - d
    Next i
    GetDaySegments = count
End Function

' The last graph-sheet column is reserved for chart-owned time-of-day values.
' Range references avoid Excel's limitations on very long series array formulas.
Private Sub CheckDailyTimeColumn(ByVal gs As Worksheet)
    Dim v As Variant, col As Long
    col = gs.Columns.Count
    v = gs.Cells(1, col).Value2
    If IsError(v) Then Fail "The graph sheet's last column is already in use."
    If CStr(v) = DAILY_TIME_MARKER Then Exit Sub
    If Application.WorksheetFunction.CountA(gs.Columns(col)) > 0 Then _
        Fail "The graph sheet's last column is already in use."
End Sub

Private Sub ClearDailyTimes(ByVal gs As Worksheet)
    Dim col As Long, last As Long, v As Variant
    col = gs.Columns.Count
    v = gs.Cells(1, col).Value2
    If IsError(v) Then Exit Sub
    If CStr(v) <> DAILY_TIME_MARKER Then Exit Sub
    last = gs.Cells(gs.Rows.Count, col).End(xlUp).Row
    gs.Range(gs.Cells(1, col), gs.Cells(last, col)).ClearContents
    gs.Columns(col).Hidden = False
End Sub

Private Sub WriteDailyTimes(ByVal gs As Worksheet, ByRef times As Variant)
    Dim col As Long
    col = gs.Columns.Count
    gs.Cells(1, col).Value2 = DAILY_TIME_MARKER
    gs.Cells(2, col).Resize(UBound(times, 1), 1).Value2 = times
    gs.Columns(col).Hidden = True
End Sub

Private Sub FinishPlot(ByVal chart As Chart, ByVal title As String, _
                       ByVal showLegend As Boolean, ByVal maxValue As Double)
    With chart
        .HasTitle = True
        .ChartTitle.Text = " "
        .ChartTitle.Format.TextFrame2.TextRange.Text = title
        .ChartTitle.Font.Size = 10
        .HasLegend = showLegend
        If .HasLegend Then .Legend.Position = xlLegendPositionBottom
        .DisplayBlanksAs = xlNotPlotted
        .PlotVisibleOnly = False
        If maxValue > 0 Then maxValue = maxValue * 1.2 Else maxValue = 1
        With .Axes(xlValue)
            .MinimumScale = 0
            .MaximumScale = maxValue
        End With
    End With
End Sub

Private Sub Set24HourAxis(ByVal chart As Chart)
    With chart.Axes(xlCategory, xlPrimary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorUnit = 1# / 24#
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormat = "[h]:mm"
        .TickLabels.Orientation = 45
        .MajorTickMark = xlTickMarkOutside
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionLow
    End With
End Sub

Private Sub BuildDailyCharts(ByVal ds As Worksheet, ByVal gs As Worksheet, _
    ByRef specs() As CounterSpec, ByVal itemCount As Long, ByVal firstRow As Long, _
    ByVal lastRow As Long, ByVal dayCount As Long, ByRef days() As Double, _
    ByRef firsts() As Long, ByRef lasts() As Long, ByVal widthPx As Double, _
    ByVal heightPx As Double, ByVal wn As Excel.Window)
    Dim i As Long, d As Long, co As ChartObject, ser As Series
    Dim dayColors() As Long
    Dim x As Range, y As Range, allValues As Range, maxValue As Double
    Dim width As Double, height As Double, gap As Double, dashStyles As Variant
    ChartSize wn, widthPx, heightPx, width, height, gap
    dashStyles = Array(msoLineSolid, msoLineDash, msoLineRoundDot, _
                       msoLineDashDot, msoLineLongDash, msoLineDashDotDot)
    ' Build one chronological date-to-color mapping shared by every chart.
    ReDim dayColors(1 To dayCount)
    For d = 1 To dayCount
        dayColors(d) = DailyLineColor(d)
    Next d
    For i = 1 To itemCount
        Application.StatusBar = "Creating daily chart " & CStr(i) & "/" & CStr(itemCount)
        Set co = gs.ChartObjects.Add(0, (i - 1) * (height + gap), width, height)
        co.Name = CHART_PREFIX & CStr(i)
        co.Placement = xlFreeFloating
        co.Chart.ChartType = xlXYScatterLinesNoMarkers
        Do While co.Chart.SeriesCollection.Count > 0
            co.Chart.SeriesCollection(1).Delete
        Loop
        For d = 1 To dayCount
            Set x = gs.Range(gs.Cells(firsts(d) - firstRow + 2, gs.Columns.Count), _
                             gs.Cells(lasts(d) - firstRow + 2, gs.Columns.Count))
            Set y = ds.Range(ds.Cells(firsts(d), specs(i).ColumnIndex), _
                             ds.Cells(lasts(d), specs(i).ColumnIndex))
            Set ser = co.Chart.SeriesCollection.NewSeries
            ser.Name = Format$(CDate(days(d)), "yyyy/mm/dd")
            ser.XValues = x
            ser.Values = y
            ser.MarkerStyle = xlMarkerStyleNone
            If firsts(d) = lasts(d) Then
                ser.MarkerStyle = xlMarkerStyleCircle
                ser.MarkerSize = 3
                ser.MarkerForegroundColor = dayColors(d)
                ser.MarkerBackgroundColor = dayColors(d)
            End If
            ser.Format.Line.Visible = msoTrue
            ser.Format.Line.ForeColor.RGB = dayColors(d)
            ser.Format.Line.Weight = 1.5
            ser.Format.Line.DashStyle = dashStyles((d - 1) Mod 6)
            ser.Smooth = False
        Next d
        Set allValues = ds.Range(ds.Cells(firstRow, specs(i).ColumnIndex), _
                                 ds.Cells(lastRow, specs(i).ColumnIndex))
        maxValue = Application.WorksheetFunction.Max(allValues)
        FinishPlot co.Chart, CounterTitle(specs(i)), True, maxValue
        Set24HourAxis co.Chart
    Next i
End Sub

' Deterministic palette by chronological date index within the selected period.
' Golden-angle hue spacing avoids a short repeating palette for long periods.
Private Function DailyLineColor(ByVal dayIndex As Long) As Long
    Const SATURATION As Double = 0.7
    Const BRIGHTNESS As Double = 0.8
    Dim hue As Double, sector As Long, fraction As Double
    Dim p As Double, q As Double, t As Double
    Dim red As Double, green As Double, blue As Double
    hue = 0.58 + (dayIndex - 1) * 0.618033988749895
    hue = (hue - Fix(hue)) * 6#
    sector = Fix(hue)
    fraction = hue - sector
    p = BRIGHTNESS * (1# - SATURATION)
    q = BRIGHTNESS * (1# - SATURATION * fraction)
    t = BRIGHTNESS * (1# - SATURATION * (1# - fraction))
    Select Case sector
        Case 0
            red = BRIGHTNESS: green = t: blue = p
        Case 1
            red = q: green = BRIGHTNESS: blue = p
        Case 2
            red = p: green = BRIGHTNESS: blue = t
        Case 3
            red = p: green = q: blue = BRIGHTNESS
        Case 4
            red = t: green = p: blue = BRIGHTNESS
        Case Else
            red = BRIGHTNESS: green = p: blue = q
    End Select
    DailyLineColor = RGB(CLng(red * 255#), CLng(green * 255#), CLng(blue * 255#))
End Function
