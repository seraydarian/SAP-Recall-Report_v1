Option Compare Database

Public Sub NewExportreCallReport()

Dim mydate As String

mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))

'acSpreadsheetTypeExcel12

DoCmd.TransferSpreadsheet acExport, 10, "ReCall-Customer_Report", "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\ReCall-Report.xlsx", True, "ReCall-Report"
'DoCmd.TransferSpreadsheet acExport, 10, "Delivery", "C:\Users\" & Environ("Username") & "\Desktop\LineFill-" & mydate & "\LineFill-Report.xlsx", True, "All Delivery"
'DoCmd.TransferSpreadsheet acExport, 10, "Ship", "C:\Users\" & Environ("Username") & "\Desktop\LineFill-" & mydate & "\LineFill-Report.xlsx", True, "All Shipments"

End Sub



Public Sub FormatExcel()

Dim xlApp As Object
Dim wb As workbook
Dim ws As Worksheet
'Dim temp As String
'Dim TempFileName As String
Dim myFileName As String
'Dim delVBAP As String
Dim mydate As String

'temp = Format(Now(), "mm-dd-yy")
'TempFileName = "VBAP-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\ReCall-Report.xlsx"

    Set xlApp = CreateObject("Excel.Application")
   ' xlApp.Application.ScreenUpdating = False
    xlApp.Visible = True
     
    Set wb = xlApp.Workbooks.Open(myFileName, 2, False)
    
    Dim i As Integer
     Dim ws_num As Integer
     
     Dim starting_ws As Worksheet
     
     Set starting_ws = ActiveSheet
     ws_num = wb.Worksheets.Count
     
     For i = 1 To ws_num
     wb.Worksheets(i).Activate
     'Do work here
     
    'insert a new row
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells.Select

    With Selection.Font
        .Name = "Comic Sans MS"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Rows("2:2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Comic Sans MS"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
     
    'Change to Title of the Report will rint on every sheet
    wb.Worksheets(i).Cells(1, 1) = 1 'this sets cell A1 as each Sheet to "1"
     
    Next
    
     wb.SaveAs "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\ReCall-Summary-Report-Done.xlsx", xlOpenXMLWorkbook, , , , , False, , xlLocalSessionChanges
    
    starting_ws.Activate 'activate the worksheet that was orginall active

    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing

End Sub

