set wsc = CreateObject("WScript.Shell")
Do
    WScript.Sleep(1*50*1000)
    wsc.SendKeys("{F13}")
Loop