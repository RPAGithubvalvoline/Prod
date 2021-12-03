On Error Resume Next
set out=WScript.CreateObject("Outlook.Application")
set mapi=out.GetNameSpace("MAPI")



For Each ac In out.Session.Accounts

'ac = "botrunner_1@valvoline.com"
Set Store = ac.Session.Stores

                For Each e In Store


If Trim(LCase(e)) = LCase("EMEAAPINVOICES/Shared/Valvoline") then
               
              Set ReadFolder = e.GetRootFolder.Folders("Inbox") 

          

 set myFolder = e.GetRootFolder.Folders("NoSubjectRecalled")



year_no = Year(Now())



'set subFolder = myFolder.Folders("NoSubjectRecalled")

Set Items = myFolder.Items
For lngCount = Items.Count To 1 Step -1
	Set m = Items(lngCount)
	m.Move(ReadFolder)
Next
End If
Next
Next

Set outlook = Nothing
Set CaseTitle = Nothing
Set session = Nothing

WScript.Quit