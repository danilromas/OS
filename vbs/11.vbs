Dim shell
Set shell = WScript.CreateObject("WScript.Shell")


Dim short, desk, deskpath

deskpath = shell.SpecialFolders("Desktop")

Set short = shell.CreateShortcut(deskpath & "\excel.lnk")

short.TargetPath = shell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE")
short.WorkingDirectory = shell.ExpandEnvironmentStrings("%windir%")
short.WindowStyle = 4
short.IconLocation = shell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE, 0")
short.Save

WScript.Echo "create"