'Tool for changing dates, adding notes, and setting a reason code - Version 5.0 12/08/2022 - daniel_buccos@cargill.com
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
'CREATE VARIABLES
Dim objExcel
Dim objSheet, intRow, i
Dim Current
Dim Current2
Dim Current3
Dim Date
Dim Text
Dim Choice
Dim ActID
Date = FormatDateTime(Now, vbShortDate)
'EXCEK HOOKS
Set objExcel = GetObject(,"Excel.Application") 
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
For i = 2 to objSheet.UsedRange.Rows.Count 
COL1 = Trim(CStr(objSheet.Cells(i, 1).Value)) 'ColumnA
COL2 = Trim(CStr(objSheet.Cells(i, 2).Value)) 'ColumnB
COL3 = Trim(CStr(objSheet.Cells(i, 3).Value)) 'ColumnC
COL4 = Trim(CStr(objSheet.Cells(i, 4).Value)) 'ColumnD
COL5 = Trim(CStr(objSheet.Cells(i, 5).Value)) 'ColumnE
'Check if col1 is blank or contains only numbers. exit loop if it does not.
if col1 = "" Then
   Exit For
End if
If Not IsNumeric(col1) Then
   Exit For
End If
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = col1
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
session.findById("wnd[0]/tbar[0]/btn[0]").press
'CLICK THROUGH POPUPS
If Session.Children.Count > 1 Then
   session.findById("wnd[1]").sendVKey 0
End If
If Session.Children.Count > 1 Then
   session.findById("wnd[1]").sendVKey 0
End If
If Session.Children.Count > 1 Then
   session.findById("wnd[1]").sendVKey 0
End If
'UPDATE DATE
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_MKAL").press
session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[3]").select
session.findById("wnd[1]/usr/ctxtRV45A-S_ETDAT").text = col2
session.findById("wnd[1]/usr/ctxtRV45A-S_ETDAT").setFocus
session.findById("wnd[1]/usr/ctxtRV45A-S_ETDAT").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[7]").press
If Session.Children.Count > 2 Then
   session.findById("wnd[2]").sendVKey 0
End If
If Session.Children.Count > 2 Then
   session.findById("wnd[2]").sendVKey 0
End If
If Session.Children.Count > 2 Then
   session.findById("wnd[2]").sendVKey 0
End If
'UPDATE NOTES
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[10]").select
'NEW CODE
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "ZH20","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "ZH20","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "ZH20","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes 112,112
Text = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/btnTP_DELETE").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = Text + vbcr + vbcr + col3 + " " + Date
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'
'SAVE ORDER
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
ActID = session.ActiveWindow.Text
If Instr(ActID, "Credit") > 0 Then
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   objExcel.Cells(i, 5).Value = "Credit Hold"
End If
'SET CHANGE REASON CODE
ActID = session.ActiveWindow.Text
If ActID = "Reason for Order change" Then
   Choice = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[" + col4 +",0]"
   session.findById(Choice).select
   session.findById(Choice).setFocus
   session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
ActID = session.ActiveWindow.Text
If Instr(ActID, "Incomplete") > 0 Then
   COL5 = Trim(CStr(objSheet.Cells(i, 5).Value)) 'ColumnE
   objExcel.Cells(i, 5).Value = col5 + " / Order saved incomplete"
   session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
   ActID = session.ActiveWindow.Text
   If ActID = "Reason for Order change" Then
      Choice = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[" + col4 +",0]"
      session.findById(Choice).select
      session.findById(Choice).setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If
End IF
next

'Hardcoded Script file name becuase I can't figure out how to get script name without Wscript
fname = "Date_Change_and_notes.vbs"

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