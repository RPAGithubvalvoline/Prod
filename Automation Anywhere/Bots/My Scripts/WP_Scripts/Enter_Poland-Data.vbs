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
session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").text = "31.01.2020"
session.findById("wnd[0]/usr/ctxtBKPF-BLART").text = "LB"
session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "0481"
session.findById("wnd[0]/usr/ctxtBKPF-BUDAT").text = "31.01.2020"
session.findById("wnd[0]/usr/txtBKPF-XBLNR").text = "-BOT"
session.findById("wnd[0]/usr/txtBKPF-MONAT").text = "4"
session.findById("wnd[0]/usr/ctxtBKPF-WAERS").text = "PLN"
session.findById("wnd[0]/usr/ctxtRF05A-KONTO").text = "114112"
session.findById("wnd[0]/usr/ctxtBSEG-VALUT").text = "09.02.2019"
session.findById("wnd[0]/usr/ctxtBSEG-PRCTR").text = "9900004"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/mbar/menu[4]/menu[2]/menu[0]").select

session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/usr/txtBKPF-XBLNR").setFocus
session.findById("wnd[0]/usr/txtBKPF-XBLNR").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press

session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
