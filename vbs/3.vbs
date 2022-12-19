set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite "HKCU\NewKey\NewValue", 1, "REG_DWORD"
WSHShell.Run "regedit",3
