Option Compare Database



Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
Dim SapGuiAuto As Variant
Dim SapApp As Variant
Dim Connection As Variant
Dim Session As Variant
Dim wscript As Variant

Sub ConnectSAP()


On Error GoTo errmsg
   '   SETS SAP OBJECTS
    If Not IsObject(SapApp) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SapApp = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = SapApp.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If
    If IsObject(wscript) Then
       wscript.ConnectObject Session, "on"
       wscript.ConnectObject Application, "on"
    End If
    
'clear the tcode in SAP
Session.findById("wnd[0]").resizeWorkingPane 222, 29, False
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
Session.findById("wnd[0]").sendVKey 0
       
Exit Sub
errmsg:
MsgBox ("Open and login to SAP first")
End
End Sub







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







Sub MSEG_SAP_Download()

Call ConnectSAP

DoCmd.OpenQuery "qry_Material"
DoCmd.RunCommand acCmdSelectAllRecords
DoCmd.RunCommand acCmdCopy
DoCmd.SetWarnings False
DoCmd.Close
DoCmd.SetWarnings True



Dim MyXLMSEG As Object    ' Variable to hold reference                              ' to Microsoft Excel.
Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim mydate As String


temp = Format(Now(), "mm-dd-yy")
TempFileName = "MSEG-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\ReCall-" & mydate & "\" & TempFileName

Session.findById("wnd[0]").Maximize

If Session.Info.Transaction <> "ZSE16N" Then
    'Temp measure, have it navgiate for the future
    Session.sendcommand ("/nZSE16N")
End If
'Material Document
Session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "MSEG"
Session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[1]/btn[18]").press
'Use Finder to select the fields    Material doc number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MBLNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]").sendVKey 0
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Material Doc Year
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MJAHR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Item in Material doc
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "ZEILE"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Movement Type
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "BWART"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Material Number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MATNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        Plant
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "WERKS"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Batch
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "CHARG"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Customer Account number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "KUNNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        Quantity
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MENGE"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True











'Load the Ssaled doc into clipboard and past them into MATNR
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MATNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press



'load the Sales doc from memory for searching
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").SetFocus
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press
Session.findById("wnd[1]/tbar[0]/btn[24]").press
Session.findById("wnd[1]/tbar[0]/btn[8]").press










DoCmd.OpenQuery "qry_Batch"
DoCmd.RunCommand acCmdSelectAllRecords
DoCmd.RunCommand acCmdCopy
DoCmd.SetWarnings False
DoCmd.Close
DoCmd.SetWarnings True





'Load the Ssaled doc into clipboard and past them into VBLEN
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "CHARG"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press



'load the Sales doc from memory for searching
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").SetFocus
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press
Session.findById("wnd[1]/tbar[0]/btn[24]").press
Session.findById("wnd[1]/tbar[0]/btn[8]").press
Session.findById("wnd[0]/tbar[1]/btn[8]").press

'save excel file and close
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&XXL"
'long wait for SAP to Selct large range
Session.findById("wnd[1]/tbar[0]/btn[0]").press
'Sleep (2000)
'Need to have a folder named
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = TempFileName
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 1
Session.findById("wnd[1]/tbar[0]/btn[0]").press

'Sleep (1000)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Allow for exter time to haev excel open
'Sleep (40000)


On Error Resume Next ' Defer error trapping.

    Set MyXLMSEG = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    
    
    Err.Clear    ' Clear Err object in case error occurred.
' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
    
' Set the object variable to reference the file you want to see.
    Set MyXLMSEG = GetObject(myFileName)
    
    
' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    MyXLMSEG.Application.Visible = True
    MyXLMSEG.Parent.Windows(1).Visible = True
    
    
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    If ExcelWasNotRunning = True Then
        MyXLMSEG.Application.Quit
    End If
    
    'MyXLVBAK.Save
    'MyXLVBAK.Close
   'ActiveWorkbook.Close

    Set MyXLMSEG = Nothing    ' Release reference to the
                                ' application and spreadsheet

TempFileName = ""
temp = ""
myFileName = ""
mydate = ""
MyXLMSEG = ""
Err.Clear


MsgBox "RecAll Report's MSEG Material Document  COMPLETE!"


End Sub


Public Sub MSEG_Loader()

Dim xlApp As Object
Dim wb As workbook
Dim ws As Worksheet
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim delMSEG As String
Dim mydate As String

temp = Format(Now(), "mm-dd-yy")
TempFileName = "MSEG-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName


    Set xlApp = CreateObject("Excel.Application")
   ' xlApp.Application.ScreenUpdating = False
    xlApp.Visible = False
    
    Set wb = xlApp.Workbooks.Open(myFileName, 2, False)
    Set ws = wb.Worksheets("Sheet1") 'EDIT TO ACTUAL SHEET NAME HERE

    delMSEG = "DELETE * From [tbl-MSEG]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL delMSEG
    DoCmd.SetWarnings True

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl-MSEG", "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName, True

    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''




Public Sub KNA1_SAP_Download()

Call ConnectSAP

DoCmd.OpenQuery "qry_Customer_Run_List"
DoCmd.RunCommand acCmdSelectAllRecords
DoCmd.RunCommand acCmdCopy
DoCmd.SetWarnings False
DoCmd.Close
DoCmd.SetWarnings True

Dim MyXLKNA As Object    ' Variable to hold reference                              ' to Microsoft Excel.
Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim mydate As String


temp = Format(Now(), "mm-dd-yy")
TempFileName = "KNA1-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName

Session.findById("wnd[0]").Maximize

If Session.Info.Transaction <> "ZSE16N" Then
    'Temp measure, have it navgiate for the future
    Session.sendcommand ("/nZSE16N")
End If

Session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "KNA1"
Session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[1]/btn[18]").press

'Use Finder to select the fields    Shipment Number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "KUNNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]").sendVKey 0
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Country Key
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "LAND1"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Use Finder to select the fields    Name 1
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "NAME1"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Use Finder to select the fields    Name 2
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "NAME2"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True



'Use Finder to select the fields    City
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "ORT01"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields   Post Office
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "PSTLZ"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        Region State
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "REGIO"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Use Finder to select the fields        House number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "STRAS"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Use Finder to select the fields        First Phone number
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "TELF1"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Use Finder to select the fields        Fax
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "TELFX"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Load the Ssaled doc into clipboard and past them into VBLEN
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "KUNNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
'load the Sales doc from memory for searching
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").SetFocus
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press
Session.findById("wnd[1]/tbar[0]/btn[24]").press
Session.findById("wnd[1]/tbar[0]/btn[8]").press
Session.findById("wnd[0]/tbar[1]/btn[8]").press
'save excel file and close
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&XXL"
'long wait for SAP to Selct large range
Session.findById("wnd[1]/tbar[0]/btn[0]").press
'Sleep (2000)
'Need to have a folder named
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = TempFileName
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 1
Session.findById("wnd[1]/tbar[0]/btn[0]").press


'Sleep (1000)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Allow for exter time to haev excel open
'Sleep (40000)

On Error Resume Next ' Defer error trapping.

    Set MyXLKNA = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    
    
    Err.Clear    ' Clear Err object in case error occurred.
' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
    
' Set the object variable to reference the file you want to see.
    Set MyXLKNA = GetObject(myFileName)
    
    
' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    MyXLKNA.Application.Visible = True
    MyXLKNA.Parent.Windows(1).Visible = True
    
    
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    If ExcelWasNotRunning = True Then
        MyXLKNA.Application.Quit
    End If
    
    'MyXLKNA.Save
    'MyXLKNA.Close
    'ActiveWorkbook.Close

    Set MyXLKNA = Nothing    ' Release reference to the
                                ' application and spreadsheet

TempFileName = ""
temp = ""
myFileName = ""
mydate = ""
MyXLKNA = ""
Err.Clear

MsgBox "Recall Report's KNA1 Sales Document Delivery Header Item COMPLETE!"



End Sub


Public Sub KNA1_Loader()
'shipment Heder table
Dim xlApp As Object
Dim wb As workbook
Dim ws As Worksheet
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim delKNA1 As String
Dim mydate As String


temp = Format(Now(), "mm-dd-yy")
TempFileName = "KNA1-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName


    Set xlApp = CreateObject("Excel.Application")
   ' xlApp.Application.ScreenUpdating = False
    xlApp.Visible = False
    
    Set wb = xlApp.Workbooks.Open(myFileName, 2, False)
    Set ws = wb.Worksheets("Sheet1") 'EDIT TO ACTUAL SHEET NAME HERE
    
  
    delKNA1 = "DELETE * From [tbl-KNA1]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL delKNA1
    DoCmd.SetWarnings True

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl-KNA1", "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName, True

    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing

'Call CloseAllExcel


End Sub







''''''''''''''''''''''




Sub MARA_SAP_Download()

Call ConnectSAP

DoCmd.OpenQuery "qry_Material"
DoCmd.RunCommand acCmdSelectAllRecords
DoCmd.RunCommand acCmdCopy
DoCmd.SetWarnings False
DoCmd.Close
DoCmd.SetWarnings True

Dim MyXLMARA As Object    ' Variable to hold reference                              ' to Microsoft Excel.
Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim mydate As String


temp = Format(Now(), "mm-dd-yy")
TempFileName = "MARA-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName

Session.findById("wnd[0]").Maximize

If Session.Info.Transaction <> "ZSE16N" Then
    'Temp measure, have it navgiate for the future
    Session.sendcommand ("/nZSE16N")
End If

Session.findById("wnd[0]/usr/ctxtGD-TAB").Text = "MARA"
Session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[1]/btn[18]").press
'Use Finder to select the fields    Material
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MATNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]").sendVKey 0
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Material Type
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MTART"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Lab Office
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "LABOR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Product Hierachy
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "PRDHA"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Batch Managed
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "XCHPF"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        X-Plant Status
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MSTAE"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    X-DChain Staus
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MSTAV"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields    Rem. Shelf Life
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MHDRZ"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        Tot. shelf life
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MHDHB"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
'Use Finder to select the fields        Material description
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MAKTX"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True

'Load the Ssaled doc into clipboard and past them into Material
Session.findById("wnd[0]/tbar[0]/btn[71]").press
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "MATNR"
Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
Session.findById("wnd[1]/tbar[0]/btn[0]").press
'load the Sales doc from memory for searching
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").SetFocus
Session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,0]").press
Session.findById("wnd[1]/tbar[0]/btn[24]").press
Session.findById("wnd[1]/tbar[0]/btn[8]").press
Session.findById("wnd[0]/tbar[1]/btn[8]").press

'save excel file and close
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&XXL"
'long wait for SAP to Selct large range
Session.findById("wnd[1]/tbar[0]/btn[0]").press
'Sleep (2000)
'Need to have a folder named
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = TempFileName
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 1
Session.findById("wnd[1]/tbar[0]/btn[0]").press




'Sleep (1000)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Allow for exter time to haev excel open
'Sleep (40000)

On Error Resume Next ' Defer error trapping.

    Set MyXLMARA = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelWasNotRunning = True
    
    
    Err.Clear    ' Clear Err object in case error occurred.
' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
    
' Set the object variable to reference the file you want to see.
    Set MyXLMARA = GetObject(myFileName)
    
    
' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    MyXLMARA.Application.Visible = True
    MyXLMARA.Parent.Windows(1).Visible = True
    
    
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    If ExcelWasNotRunning = True Then
        MyXLMARA.Application.Quit
    End If
    
    'MyXLMARA.Save
    'MyXLMARA.Close
    'ActiveWorkbook.Close

    Set MyXLMARA = Nothing    ' Release reference to the
                                ' application and spreadsheet


TempFileName = ""
temp = ""
myFileName = ""
mydate = ""
MyXLMARA = ""
Err.Clear

MsgBox "Recall Report's MARA Sales Document Header COMPLETE!"


End Sub


Public Sub MARA_Loader()

Dim xlApp As Object
Dim wb As workbook
Dim ws As Worksheet
Dim temp As String
Dim TempFileName As String
Dim myFileName As String
Dim delMARA As String
Dim mydate As String

temp = Format(Now(), "mm-dd-yy")
TempFileName = "MARA-" & temp & ".XLSX"
mydate = Trim(VBA.Format(Now(), "MM-DD-YY"))
myFileName = "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName


    Set xlApp = CreateObject("Excel.Application")
   ' xlApp.Application.ScreenUpdating = False
    xlApp.Visible = False
    
    Set wb = xlApp.Workbooks.Open(myFileName, 2, False)
    Set ws = wb.Worksheets("Sheet1") 'EDIT TO ACTUAL SHEET NAME HERE

    delMARA = "DELETE * From [tbl-MARA]"

    DoCmd.SetWarnings False
    DoCmd.RunSQL delMARA
    DoCmd.SetWarnings True

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl-MARA", "C:\Users\" & Environ("Username") & "\Desktop\Build-Recall-Report\Recall-" & mydate & "\" & TempFileName, True

    wb.Save
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "Recall Report's HAS COMPLETED LOADING ALL THE TABLES!!!! COMPLETE!"

End Sub


