Set outlook = createobject("outlook.application")
Set session = outlook.getnamespace("mapi")
session.logon
Set inbox = session.getdefaultfolder(6)
Dim DelXl
DelXl = 0
Set Items = Inbox.Items
For lngCount = Items.Count To 1 Step -1
	Set m = Items(lngCount)
	If m.unread Then
	Set m = Items(lngCount)
	If m.MessageClass = "IPM.Outlook.Recall" Then
		subject = m.subject
		DelXl = 1
	End If
	End If
Next 
	
session.logoff
Set outlook = Nothing
Set session = Nothing
wscript.stdout.writeline DelXl
WScript.Quit