On Error Resume Next
Dim Args1,Args2,out,mapi,myFolder,subFolder,MyNextLevelFolder,subFolder1,myNewFolder,currMonth,mnthname
Args1=WScript.Arguments.Item(0)
Args2=WScript.Arguments.Item(1)
'Args1="Invoices-2019"
'Args2="02-Sep"
curMonth=right("00" & month(now),2)
mnthname = left(MonthName(curMonth),3)
Args2=curMonth & "-" & mnthname
set out=WScript.CreateObject("Outlook.Application")
set mapi=out.GetNameSpace("MAPI")

'NEW CODE START

For Each ac In out.Session.Accounts
'msgbox ac
'ac = "botrunner_1@valvoline.com"
Set Store = ac.Session.Stores

                For Each e In Store
If Trim(LCase(e)) = LCase("Valv Vendorinvoice/Shared/Valvoline") then
                'MsgBox e 
              Set myFolder = e.GetRootFolder.Folders("Valvoline")

'NEW CODE END

'set myFolder = MAPI.Folders("horesh.sharma@valvoline.com").Folders("Valvoline_NAM")
'Code for Creating Invoice Folder as per current year
set subFolder = myFolder.Folders(Args1)
On Error Resume Next
if subFolder Is Nothing Then
subFolder = myFolder.Folders.Add(Args1)
End If
'Code for Creating Month wise folder in Invoice Folder
On Error Resume Next 
 set MyNextLevelFolder = MyFolder.Folders(Args1)
set subFolder1 = MyNextLevelFolder.Folders(Args2)
if subFolder1 Is Nothing Then
set myNewFolder = MyNextLevelFolder.Folders.Add(Args2)
End If
Wscript.Stdout.writeline Args2
set  mynextlevelfolder = nothing
set myNewFolder=Nothing
set Args1=Nothing
set Args2=Nothing
set out=Nothing
set myFolder=Nothing
set subFolder=Nothing
set MyNextLevelFolder=Nothing
set subFolder1=Nothing
exit for
End if
               next
next 

WScript.Quit