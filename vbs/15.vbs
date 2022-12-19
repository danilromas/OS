Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

Set colAccounts = GetObject("WinNT://" & strComputer & "")
Set objUser = colAccounts.Create("user", "LocalAdmin")
objUser.SetPassword "Hello123456789"
objUser.SetInfo

Set objGroup = GetObject("WinNT://" & strComputer & "/Администраторы,group")
Set objUser = GetObject("WinNT://" & strComputer & "/LocalAdmin,user")
objGroup.Add(objUser.ADsPath)

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
UserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = UserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag
objUser.SetInfo