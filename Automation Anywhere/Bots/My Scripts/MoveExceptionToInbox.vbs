Dim oApp
Dim oMapi
Dim oMail
year_no = Year(Now())
Args1="Invoices-"&year_no
curMonth=right("00" & month(now),2)
mnthname = left(MonthName(curMonth),3)
Args2=curMonth & "-" & mnthname

Set oApp = GetObject(, "OUTLOOK.APPLICATION")
Set objExcel = CreateObject("Excel.Application")

objExcel.visible = True
CsvPath = WScript.Arguments.Item(0)
Set objWorkbook = objExcel.Workbooks.Open(CsvPath)

If (oApp Is Nothing) Then
	Set oApp = CreateObject("OUTLOOK.APPLICATION")
End If

Set oMapi = CreateObject("Outlook.Application")
Set nameSpace = oMapi.GetNamespace("MAPI")

Set Inbox1 = nameSpace.Folders("daksh.srivastava@valvoline.com").Folders("Inbox")
Set MyFolders = nameSpace.Folders("daksh.srivastava@valvoline.com").Folders("Valvoline")
Set MyFolders = MyFolders.Folders(Args1)

Set SubFolder = MyFolders.Folders(Args2)


x = 2

Do Until objExcel.Cells(x,1).Value = ""
	if objExcel.Cells(x,6).Value <> "" and objExcel.Cells(x,6).Value <> "No Attachment" and objExcel.Cells(x,6).Value <> "Moved to Statements" then
		vSubject = objExcel.Cells(x,7).Value
		Msgbox vSubject
		vReceivedTime = objExcel.Cells(x,3).Value
		vReceived =  Mid(vReceivedTime, 2)
		Msgbox vReceived
		Set Items = SubFolder.Items
		For lngCount = Items.Count To 1 Step -1
			Set item1 = Items(lngCount)
			if item1.Subject = vSubject  then
				Recieved_Time = item1.ReceivedTime
				Recieved_Time = DateAdd("s",-2,Recieved_Time)
				Msgbox Recieved_Time
				if Recieved_Time = vReceived then
				Msgbox "In"
					item1.unread = true
					item1.move Inbox1
				end if
			end if
		next
	end if
	x = x + 1
Loop

objWorkbook.close
objExcel.Quit

Set oApp = Nothing
Set oMapi = Nothing
Set oMail = Nothing
Set oHTML = Nothing
Set oElColl = Nothing