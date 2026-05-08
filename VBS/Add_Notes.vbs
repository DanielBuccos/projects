'Tool for adding notes to internal sales order text - Version 2.0 4/27/2023 - daniel_buccos@cargill.com
'T-Code VA02
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
Dim objExcel
Dim objSheet, intRow, i
Dim Current
Dim Date
Dim String
Dim Field
Dim Field2
Dim Text
Dim ActID
Date = FormatDateTime(Now, vbShortDate)
Set objExcel = GetObject(,"Excel.Application") 
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
For i = 2 to objSheet.UsedRange.Rows.Count 
COL1 = Trim(CStr(objSheet.Cells(i, 1).Value)) 'ColumnA
COL2 = Trim(CStr(objSheet.Cells(i, 2).Value)) 'ColumnB
COL3 = Trim(CStr(objSheet.Cells(i, 3).Value)) 'ColumnC
COL4 = Trim(CStr(objSheet.Cells(i, 4).Value)) 'ColumnD
COL5 = Trim(CStr(objSheet.Cells(i, 5).Value)) 'ColumnE
COL6 = Trim(CStr(objSheet.Cells(i, 6).Value)) 'ColumnF
COL7 = Trim(CStr(objSheet.Cells(i, 7).Value)) 'ColumnG
COL8 = Trim(CStr(objSheet.Cells(i, 8).Value)) 'ColumnH
COL9 = Trim(CStr(objSheet.Cells(i, 9).Value)) 'ColumnI
COL10 = Trim(CStr(objSheet.Cells(i, 10).Value)) 'ColumnJ
'Check if col1 is blank or contains only numbers. exit loop if it does not.
if col1 = "" Then
    Exit For
End if
If Not IsNumeric(col1) Then
    Exit For
End If
'Begin SAP code
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = col1
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
session.findById("wnd[0]/tbar[0]/btn[0]").press
If Session.Children.Count > 1 Then
    session.findById("wnd[1]").sendVKey 0
End If
If Session.Children.Count > 1 Then
    session.findById("wnd[1]").sendVKey 0
End If
If Session.Children.Count > 1 Then
    session.findById("wnd[1]").sendVKey 0
End If
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[10]").select
'CREATE STRING
Field = objExcel.Cells(1, 3).Value
if Field = "Please choose a text field to update here" Then
    MsgBox("Please choose select a text field using the drop down box in row 1, column C of the spreadsheet. Press 'OK' on this message box once completed")
End if 
Field = objExcel.Cells(1, 3).Value
if Field = "Internal Invoice" Then
    String = "ZH08"
End IF
if Field = "Loading Instructions" Then
    String = "ZH10"
End IF
if Field = "Shipping Instructions" Then
    String = "ZH11"
End IF
if Field = "Bill of Lading" Then
    String = "ZH15"
End IF
if Field = "Internal Sales Order" Then
    String = "ZH20"
End IF
if Field = "WINSHUTTLE TEXT" Then
    String = "ZH32"
End IF
'CHECK if DELAYED DELIVERY ORDER
Field2 = objExcel.Cells(1, 4).Value
if Field2 = "Add 'Delivered Delays -  ' in front of text" Then
    col2 = "Delivered Delays - " + col2
End if
'ENTER NOTES
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem String,"Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem String,"Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem String,"Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 112,112
Text = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/btnTP_DELETE").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = Text + vbcr + vbcr + col2 + " " + Date
if Field = "WINSHUTTLE TEXT" Then
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = col2
End If
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'SAVE ORDER
ActID = ""
ActID = session.ActiveWindow.Text
If Instr(ActID, "Credit") > 0 Then
   session.findById("wnd[1]").sendVKey 0
End IF
If Session.Children.Count > 1 Then
    session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
End If
next

'Hardcoded Script file name becuase I can't figure out how to get script name without Wscript
fname = "Add_Notes.vbs"

'Create log file and log variables.
Public strFileName
Public fname
Set objShell = CreateObject("WScript.Shell")
un = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
un = Replace(un, "C:\Users\", "")
filestring = "C:\Users\" & un & "\Cargill Inc\Salt Customerservice - Shared Scripts\F&WQ Scripts\Development Files\Log.vbs"
strFileName =  Chr(34) & filestring & Chr(34)
Set WshShell = CreateObject("WScript.Shell")
WSHShell.Run "wscript " & strFileName & " " & un & " " & filestring & " " & fname

MsgBox(" <('.'<) Process Completed! (>'.')>")