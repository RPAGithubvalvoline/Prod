On Error Resume Next
set out=WScript.CreateObject("Outlook.Application")
set mapi=out.GetNameSpace("MAPI")

For Each ac In out.Session.Accounts

'ac = "botrunner_1@valvoline.com"
Set Store = ac.Session.Stores

                For Each e In Store


If Trim(LCase(e)) = LCase("Valv Vendorinvoice/Shared/Valvoline") then
               
              Set ReadFolder = e.GetRootFolder.Folders("Inbox") 

             'Set ReadFolder = session.getdefaultfolder(6)

 set myFolder = e.GetRootFolder.Folders("Valvoline")



set subFolder = myFolder.Folders("NoSubjectRecalled")
On Error Resume Next
if subFolder Is Nothing Then
subFolder = myFolder.Folders.Add("NoSubjectRecalled")
End If

Csv_Path=WScript.Arguments.Item(0)

DIM fso    
Set fso = CreateObject("Scripting.FileSystemObject")
dim objExcel, objWorkbook ,headers,xlsht,i
If not(fso.FileExists(Csv_Path)) Then
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	headers = Array("SenderEmailid", "MailReceviedTime", "No of attachments", "Email Body" , "Destination Folder")
	For Each xlSht In objWorkbook.Sheets
		With xlSht
			.Rows(1).Value = "" 'This will clear out row 1
			For i = LBound(headers) To UBound(headers)
				.Cells(1, 1 + i).Value = headers(i)
			Next
			.Rows(1).Font.Bold = True
		End With
	Next
	objWorkbook.SaveAs(Csv_Path)
	objExcel.Quit
End If

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.DisplayAlerts = false   
Set objWorkbook = objExcel.Workbooks.Open(Csv_Path)
Set objSrcWorksheet = objWorkbook.Worksheets(1)
Dim j 
j = 2
Dim DelXl
DelXl = 0
Set Items = ReadFolder.Items
For lngCount = Items.Count To 1 Step -1
	Set m = Items(lngCount)
	'If m.unread Then
                
		Email_From = m.SenderEmailAddress
		Recieved_Time = m.ReceivedTime
                
		subject = m.subject
		Email_Body = m.body
		AtchCount = m.Attachments.Count
                MsgClass = m.MessageClass
		
		If (subject="") OR InStr(1, subject,"Recall",1) <> 0 OR InStr(1, subject,"Messages on hold",1) <> 0 OR MsgClass = "IPM.Note.SMIME.MultipartSigned" OR MsgClass = "IPM.Schedule.Meeting.Request" Then
			DelXl = 1
                        'MsgClass = m.MessageClass
			m.move subFolder
 			
			objSrcWorksheet.Range("A"&j&"").Value = Email_From
			objSrcWorksheet.Range("B"&j&"").Value = Recieved_Time
			objSrcWorksheet.Range("C"&j&"").Value = AtchCount
			'objSrcWorksheet.Range("D"&j&"").Value = Email_Body
			objSrcWorksheet.Range("E"&j&"").Value = "NoSubjectRecalled"
			j=j+1		
		End If
			
	'End If
Next 
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
SET objExcel = nothing	
session.logoff
Set outlook = Nothing
Set CaseTitle = Nothing
Set session = Nothing

If DelXl = 0 Then
       
	fso.DeleteFile(Csv_Path)
        
End If

WScript.Quit
exit for
End if
              
Next
Next
