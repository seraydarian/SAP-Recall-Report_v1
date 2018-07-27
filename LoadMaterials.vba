Option Compare Database

Public Sub LoadMaterials()

Dim xlApp As Object
Dim wb As workbook
Dim ws As Worksheet
'Dim temp As String
'Dim TempFileName As String
Dim myFileName As String
Dim delMaterialLoader As String
Dim mydate As String

'temp = Format(Now(), "mm-dd-yy")
'TempFileName = "MARA-" & temp & ".XLSX"
'mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Material-Loader.XLSX"


    Set xlApp = CreateObject("Excel.Application")
   ' xlApp.Application.ScreenUpdating = False
    xlApp.Visible = False
    
    Set wb = xlApp.Workbooks.Open(myFileName, 2, False)
    Set ws = wb.Worksheets("Material-Loader") 'EDIT TO ACTUAL SHEET NAME HERE

    delMaterialLoader = "DELETE * From [Material-Loader]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL delMaterialLoader
    DoCmd.SetWarnings True

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Material-Loader", "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Material-Loader.XLSX", True

    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "ReCall Report's HAS COMPLETED LOADING Material LOADER TABLES!!!! "




End Sub
