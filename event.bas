Option Explicit

Private Sub Workbook_WindowResize(ByVal Wn As Excel.Window)
    ResizePerfmonCharts Wn
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)
    ResizePerfmonCharts Wn
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Application.ActiveWindow Is Nothing Then Exit Sub
    ResizePerfmonCharts Application.ActiveWindow
End Sub