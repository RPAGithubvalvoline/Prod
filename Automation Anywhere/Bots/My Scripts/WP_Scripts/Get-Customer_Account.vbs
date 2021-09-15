On Error Resume Next

Dim Account_Code
Dim Error_Text
Dim c_code
Dim T_Cust_Acc
Dim c_name

c_code = (WScript.Arguments(0))
'c_code = "0271"
c_name = (WScript.Arguments(1))
'c_name = "*SEALMECH*"

'msgbox("check values inside script COMPANY CODE AND CUSTOMER NAME")
'MsgBox(c_code)
'MsgBox(c_name)

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
session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").caretPosition = 6
session.findById("wnd[0]").sendVKey 4


'MsgBox( "check value set before" )
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB002/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").text = c_name
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB002/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]").setFocus
'MSGBOX c_code
c_code = replace(c_code,",","")
'msgbox c_code
session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB002/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]").text = c_code

'MsgBox( "check value set after" )

'MsgBox( "check value set after DIRECT ARGUMENT" )

session.findById("wnd[1]/tbar[0]/btn[0]").press

Error_Text = session.findById("wnd[0]/sbar").Text
'MsgBox("Error_Text after entering values") 
'MsgBox( Error_Text )

If Error_Text <> "" Then
  Account_Code=""
  T_Cust_Acc = 12
 WScript.StdOut.WriteLine Account_Code
' MsgBox ("Account_Code when error " ) 
' MsgBox( Account_Code )
 ' MsgBox ("TEMP Account_Code when error " ) 
 'MsgBox( T_Cust_Acc  )
 session.findById("wnd[1]/tbar[0]/btn[12]").press
 WScript.Quit

End If

session.findById("wnd[1]").maximize
session.findById("wnd[1]/usr/lbl[79,3]").setFocus
session.findById("wnd[1]/usr/lbl[79,3]").caretPosition = 11
Account_Code = session.findById("wnd[1]/usr/lbl[79,3]").Text

session.findById("wnd[1]/usr/lbl[79,4]").setFocus
session.findById("wnd[1]/usr/lbl[79,4]").caretPosition = 11
T_Cust_Acc  = session.findById("wnd[1]/usr/lbl[79,4]").Text

'MsgBox( "Account_Code when no error")
'MsgBox( Account_Code )

'MsgBox( "Temp customer account")
'MsgBox( T_Cust_Acc   )

'MsgBox( "value written check, BEFORE" )

'WScript.StdOut.WriteLine Account_Code
'WScript.StdOut.WriteLine T_Cust_Acc  

'MsgBox( "value written check, after" )

If T_Cust_Acc = "" Then
WScript.StdOut.WriteLine Account_Code
else
Account_Code=""
WScript.StdOut.WriteLine Account_Code
End If

'MsgBox( "second Account Code ")
'MsgBox( T_Cust_Acc )
'MsgBox( "Account_Code final")
'MsgBox( Account_Code )
 
session.findById("wnd[1]").sendVKey 2

If Err.Number <> 0 Then
 Account_Code=""
 WScript.StdOut.WriteLine Account_Code
 WScript.Quit
End If
 



