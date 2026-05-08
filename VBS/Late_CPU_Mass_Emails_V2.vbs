'Batch Mass Mailing for CPU's - Version 2.1 9/19/2025 - daniel_buccos@cargill.com
'Scipt for mass mailings from CPU late pickup report
Dim objExcel
Dim objSheet, intRow, i
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
COL11 = Trim(CStr(objSheet.Cells(i, 11).Value)) 'ColumnK
COL12 = Trim(CStr(objSheet.Cells(i, 12).Value)) 'ColumnL
COL13 = Trim(CStr(objSheet.Cells(i, 13).Value)) 'ColumnM
COL14 = Trim(CStr(objSheet.Cells(i, 14).Value)) 'ColumnN
COL15 = Trim(CStr(objSheet.Cells(i, 15).Value)) 'ColumnO
COL16 = Trim(CStr(objSheet.Cells(i, 16).Value)) 'ColumnP
COL17 = Trim(CStr(objSheet.Cells(i, 17).Value)) 'ColumnQ
COL18 = Trim(CStr(objSheet.Cells(i, 18).Value)) 'ColumnR
COL19 = Trim(CStr(objSheet.Cells(i, 19).Value)) 'ColumnS
COL20 = Trim(CStr(objSheet.Cells(i, 20).Value)) 'ColumnT
COL21 = Trim(CStr(objSheet.Cells(i, 21).Value)) 'ColumnU
COL22 = Trim(CStr(objSheet.Cells(i, 22).Value)) 'ColumnV
COL23 = Trim(CStr(objSheet.Cells(i, 23).Value)) 'ColumnW
COL24 = Trim(CStr(objSheet.Cells(i, 24).Value)) 'ColumnX
COL25 = Trim(CStr(objSheet.Cells(i, 25).Value)) 'ColumnY
COL26 = Trim(CStr(objSheet.Cells(i, 26).Value)) 'ColumnZ
COL27 = Trim(CStr(objSheet.Cells(i, 27).Value)) 'ColumnAA
COL28 = Trim(CStr(objSheet.Cells(i, 28).Value)) 'ColumnAB
COL29 = Trim(CStr(objSheet.Cells(i, 29).Value)) 'ColumnAC
COL30 = Trim(CStr(objSheet.Cells(i, 30).Value)) 'ColumnAD
COL31 = Trim(CStr(objSheet.Cells(i, 31).Value)) 'ColumnAE
COL32 = Trim(CStr(objSheet.Cells(i, 32).Value)) 'ColumnAF
COL33 = Trim(CStr(objSheet.Cells(i, 33).Value)) 'ColumnAG
COL34 = Trim(CStr(objSheet.Cells(i, 34).Value)) 'ColumnAH
COL35 = Trim(CStr(objSheet.Cells(i, 35).Value)) 'ColumnAI
COL36 = Trim(CStr(objSheet.Cells(i, 36).Value)) 'ColumnAJ
'If Condition to stop from running over blank ship-to's
if not col5 = "" Then
    if col5 = "" Then
        Exit For
    End if
    'Set Outlook objects
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    objMail.Display
    objMail.Recipients.Add (col5)
    'Set Subject
    objMail.Subject = "URGENT: CARGILL SALT PAST DUE ORDER POSSIBLE CANCELLATION"
    'Set email message
    objMail.HTMLBody = "<html><body>" & _
    "Hello,<br><br>" & _
    "We currently show that you are past due in picking up the following order(s) <br><br>" & _
    "Ship to name(s): " & col4 & "<br><br>" & _
    "PO(s): " & col3 & "<br><br>" & _
    "Sales order(s) " & col2 & "<br><br>" & _
    "<b>If the site requires an appointment</b> and you have already rescheduled for " & col6 & " - no response is required.<br><br>" & _
    "Cargill Salt allows a <b>one-time</b> pickup date change; we will need to hear from you today.<br><br>" & _
    "Please respond by <b>4:00 PM CT today</b> if you need to update your pickup date. If we don't hear from you or the rescheduled pickup is missed, the order will be cancelled the next business day.<br><br>" & _
    "Please contact Cargill Salt Customer Care at Salt_Customerservice@cargill.com to communicate an update if needed.<br><br>" & _
    "Thank you.<br><br>" & _
    "Cargill Salt Customer Care<br>" & _
    "Cargill Salt<br>" & _
    "Nourishing, Enhancing, and Protecting Lives<br>" & _
    "Phone: 800-377-1017 | Salt_Customerservice@cargill.com<br>" & _
    "Cargill Salt Store Available 24/7. Ask us how you can manage your account online at <a href='https://www.cargillsaltstore.com'>www.cargillsaltstore.com</a><br><br>" & _
    "<a href='https://www.cargill.com/doc/1432095426022/cargill-salt-terms-and-conditions.pdf'>Terms & Conditions</a>" & _
    "</body></html>"
    'Delay to load in Outlook contacts
    Dim dteWait
    dteWait = DateAdd("s", 15, Now())
    Do Until (Now() > dteWait)
    Loop
    objMail.Send
    Set objMail = Nothing
    Set objOutlook = Nothing
End if
Next

'Hardcoded Script file name becuase I can't figure out how to get script name without Wscript
fname = "Late_CPU_Mass_Emails_V2.vbs"

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