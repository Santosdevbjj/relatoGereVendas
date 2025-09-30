
Attribute VB_Name = "DashboardMacros"
' Macros to toggle chart visibility and navigate between sheets.
Sub ShowChartA()
    ' In Excel UI: use Selection Pane to set chart names Chart 1, Chart 2, etc.
    ' This macro hides Chart 2 and shows Chart 1 on Dashboard sheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error Resume Next
    ws.ChartObjects("Chart 1").Visible = True
    ws.ChartObjects("Chart 2").Visible = False
End Sub

Sub ShowChartB()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error Resume Next
    ws.ChartObjects("Chart 1").Visible = False
    ws.ChartObjects("Chart 2").Visible = True
End Sub

Sub GoToPage2()
    ThisWorkbook.Sheets("Page2_Details").Activate
End Sub

Sub GoToDashboard()
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub
