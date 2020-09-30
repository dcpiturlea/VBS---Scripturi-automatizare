Option Explicit
Dim objShell, intShutdown
Dim strShutdown, strAbort
Dim sInput
sInput = InputBox("In cate minute vrei sa se opreasca PC-ul")
sInput = sInput * 60

' -s = shutdown, -t 120 = 2 minutes, -f = force programs to close
strShutdown = "shutdown.exe -s -t " & sInput & " -f"
set objShell = CreateObject("WScript.Shell")
objShell.Run strShutdown, 0, false

'go to sleep so message box appears on top
WScript.Sleep 100


Wscript.Quit