set WshShell1 = wscript.createobject("Wscript.shell")
Razdel=WshShell1.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\BITS\")
WScript.Echo Razdel