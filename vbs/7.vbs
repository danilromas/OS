Dim WshShell, Path
Set WshShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next

Path = "notepad C:\AUTOEXEC.BAT"
WshShell.Run Path

Path = "notepad C:\CONFIG.SYS"
WshShell.Run Path