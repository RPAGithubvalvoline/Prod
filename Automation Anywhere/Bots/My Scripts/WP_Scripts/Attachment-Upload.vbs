On Error Resume Next
Dim Attachment_Path
Dim File_Name

Attachment_Path = (WScript.Arguments(0))

File_Name = (WScript.Arguments(1))

'msgbox(Attachment_Path)
'msgbox(File_Name)

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
session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
'msgbox(Attachment_Path)
'msgbox(File_Name)

session.findById("wnd[1]/usr/ctxtDY_PATH").text = Replace(Attachment_Path,",","")

session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = File_Name
'msgbox("after entering values")

'msgbox(Attachment_Path)
'msgbox(File_Name)

session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2

session.findById("wnd[1]/tbar[0]/btn[0]").press
