Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
ScriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)

PYTHONW = "C:\Users\huoli\AppData\Local\Programs\Python\Python312\pythonw.exe"

WshShell.CurrentDirectory = ScriptDir

' Firewall
WshShell.Run "netsh advfirewall firewall add rule name=""yuanqi-8899"" dir=in action=allow protocol=TCP localport=8899", 0, True

' Disable sleep
WshShell.Run "powercfg -change -standby-timeout-ac 0", 0, True
WshShell.Run "powercfg -change -hibernate-timeout-ac 0", 0, True
WshShell.Run "powercfg /hibernate off", 0, True

' Watchdog (pythonw, zero window) — 由 watchdog 统一管理 launcher 启动
WshShell.Run """" & PYTHONW & """ """ & ScriptDir & "\watchdog.py""", 0, False
