Option Explicit
Dim objShell, intShutdown
Dim strDllPath, strAbort, strDllName

strDllPath = InputBox("Introduceti locatia(folderul) unde se afla aplicatia pe care vrei sa o porniti:")
strDllName = InputBox("Introduceti numele aplicatiei(DLL FILE), inclusiv extensia")
strDllName = "dotnet " & strDllName 

'strDllPath= "C:\Users\+USERNAME+\FOLDER\"
'strDllName = "dotnet ConsoleApp2.dll"

set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory  = strDllPath
objShell.Run strDllName 


Wscript.Quit