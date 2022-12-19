Set shell = WScript.CreateObject("WScript.Shell")
v = shell.ExpandEnvironmentStrings("%SystemRoot%")
MsgBox v