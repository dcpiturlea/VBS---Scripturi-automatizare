set wsc  = CreateObject("WScript.Shell")
Do
WScript.Sleep(60*1000)
wsc.SendKeys("{NUMLOCK}")
WScript.Sleep(0.01*1000)
wsc.SendKeys("{NUMLOCK}")
Loop