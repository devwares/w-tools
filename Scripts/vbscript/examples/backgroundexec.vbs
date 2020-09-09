Dim WinScriptHost
Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run "sync.bat sync.txt sync.log", 0, 1
Set WinScriptHost = Nothing
