'Batch PO Processing Tool - Version 5.0 6/1/2022 - daniel_buccos@cargill.com
'Tool for entering Sales orders into SAP PW1 using VA01 TCode
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
'Assign Vairables/Excel Hooks
Dim objExcel
Dim objSheet, intRow, i
Dim Materials
Dim Quantities
Dim Y
Dim ActID
Dim Date
Dim Delivery
Dim Difference
Date = FormatDateTime(Now, vbShortDate)
Set objExcel = GetObject(,"Excel.Application") 
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
For i = 5 to objSheet.UsedRange.Rows.Count 
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
   COL11 = Trim(CStr(objSheet.Cells(i, 11).Value)) 'ColumnK
   COL12 = Trim(CStr(objSheet.Cells(i, 12).Value)) 'ColumnL
'Check if col1 is blank or contains only numbers. exit loop if it does not.
if col1 = "" Then
   Exit For
End if
If Not IsNumeric(col1) Then
   Exit For
End If
'Enter Order Data type
session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "zor"
session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = objExcel.Cells(3, 8).Value
session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = objExcel.Cells(2, 10).Value
session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = objExcel.Cells(3, 10).Value
session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = objExcel.Cells(2, 12).Value
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = objExcel.Cells(3, 12).Value
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[8]").press
'Enter Date and Contract Number
session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON").select
session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-VBELN").text = col1
session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-KETDAT").text = col2
session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-KETDAT").setFocus
session.findById("wnd[1]/usr/tabsMYTABSTRIP/tabpRKON/ssubSUB1:SAPLV45C:0302/ctxtLV45C-KETDAT").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[7]").press
If Session.Children.Count > 2 Then
   session.findById("wnd[2]").sendVKey 0
End If
session.findById("wnd[1]/tbar[0]/btn[7]").press
If Session.Children.Count > 2 Then
   session.findById("wnd[2]").sendVKey 0
End If
session.findById("wnd[0]/usr/btnMKLO").press
'Enter Material Number and Quantity. Values are seperated by a " " into arrays
Materials = col6
Quantities = col7
Loops = x
Quantities = Split(Quantities)
Materials = Split(Materials)
Counter = 0
'Loop for number of Materials using above arrays
for each x in Materials
   session.findById("wnd[0]/mbar/menu[1]/menu[6]").select
   session.findById("wnd[1]/usr/ctxtRV45A-PO_MATNR").text = Materials(Counter)
   session.findById("wnd[1]/usr/ctxtRV45A-PO_MATNR").setFocus
   session.findById("wnd[1]/usr/ctxtRV45A-PO_MATNR").caretPosition = 9
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   session.findById("wnd[0]/usr/tblSAPMV45ATCTRL_KOPIEREN/chkRV45A-VBAP_SELKZ[0,0]").selected = true
   session.findById("wnd[0]/usr/tblSAPMV45ATCTRL_KOPIEREN/txtVBAPD-ORFMNG[5,0]").text = Quantities(Counter)
   session.findById("wnd[0]/usr/tblSAPMV45ATCTRL_KOPIEREN").columns.elementAt(8).width = 40
for z = 1 to 5
session.findById("wnd[0]/tbar[0]/btn[0]").press
Next
Counter = Counter + 1
next
ActID = "0"
session.findById("wnd[0]/tbar[1]/btn[5]").press
'Click through popup notes and check for No goods
for x = 1 To 5
   If Session.Children.Count > 1 Then
      ActID = session.ActiveWindow.Text
      If Instr(ActID, "Information") > 0 Then
         Y = session.findById("wnd[1]/usr/txtMESSTXT1").Text
         objExcel.Cells(i, 9).Value = "Potentially No Good accepted on date"
      End If
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If
next
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = col5
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 6
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[11]").select
Y = session.findById("wnd[0]/sbar").Text
'If Purchase order number already exsists, continue and notify
COL9 = objExcel.Cells(i, 9).Value 'ColumnH
If InStr(Y, "already") > 0 Then
   Y = "Duplicate PO" + " / " & col9
   objExcel.Cells(i, 9).Value = Y
End If
session.findById("wnd[0]/tbar[0]/btn[0]").press
'Enter Header data
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4351/ctxtVBKD-BSARK").text = objExcel.Cells(3, 2).Value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-BNAME").text = objExcel.Cells(3, 5).Value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-BNAME").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10/ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-BNAME").caretPosition = 14
'If Requested Delivery Time
If Instr(col3, ":") > 0 or Instr(col4, ":") > 0 Then
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").select
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_PAL").select
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC").select
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/btnCHANGEALL").press
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/ctxtVBAK-ZZCREQTIMFR").text = col3
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/ctxtVBAK-ZZCREQTIMTO").text = col4
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/ctxtVBAK-ZZCREQTIMTO").setFocus
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/ctxtVBAK-ZZCREQTIMTO").caretPosition = 8
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsGV_TABSTRIP_DTC/tabpTAB_DTC/ssubSUB_DTC:ZCOOMM_DTC:9000/btnCHANGE_SELECTED").press
End if
session.findById("wnd[0]/tbar[0]/btn[3]").press
'If Yes to preship add to preship block
if objExcel.Cells(2, 8).Value = "Yes" Then
   session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key = "51"
End If
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
'Enter through credit window if a credit hold
ActID = session.ActiveWindow.Text
If Instr(ActID, "Credit") > 0 Then
   session.findById("wnd[1]").sendVKey 0
   COL9 = Trim(CStr(objSheet.Cells(i, 9).Value)) 'ColumnG
   Y = "Credit Hold" + " / " & col9
   objExcel.Cells(i, 9).Value = Y
End if
'Place on a pricing block if Sales order is incomplete
ActID = session.ActiveWindow.Text
If Instr(ActID, "Incomplete") > 0 Then
   Difference = DateDiff("d",Date,col2)
   if Difference =< 10 Then
      session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK").key = "03"
      session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-FAKSK").setFocus
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
      COL9 = objExcel.Cells(i, 9).Value 'ColumnH
      Y = "Placed on a PI Block" + " / " & col9
      objExcel.Cells(i, 9).Value = Y
   End if
   session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
   COL9 = objExcel.Cells(i, 9).Value 'ColumnH
   Y = "Sales order incomplete" + " / " & col9
   objExcel.Cells(i, 9).Value = Y
   'Notate Credit Hold
   ActID = session.ActiveWindow.Text
   If Instr(ActID, "Credit") > 0 Then
      session.findById("wnd[1]").sendVKey 0
      COL9 = objExcel.Cells(i, 9).Value 'ColumnH
      Y = "Credit Hold" + " / " & col9
      objExcel.Cells(i, 9).Value = Y
   End if
End if
'Grab sales order number
Y = session.findById("wnd[0]/sbar").Text
Y = Replace(Y, " has been saved","",16)
objExcel.Cells(i, 8).Value = Y
Next

'Hardcoded Script file name becuase I can't figure out how to get script name without Wscript
fname = "Batch_PO_Processing_Tool.vbs"

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