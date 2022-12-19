Dim shell
Dim txt
Set shell = WScript.CreateObject("WScript.Shell")
WScript.Echo "notepad opens. . . "
txt=inputbox("new name txt file",,"new.html")
shell.Run "notepad "&txt
WScript.Sleep 1000
shell.AppActivate "notepad"