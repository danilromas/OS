# OS

'Laba 4

'1. Создать txt file

'Set fs = CreateObject("Scripting.FileSystemObject")
'Set a = fs.CreateTextFile("D:\testfile.txt", True)
'a.WriteLine("This is a test.")
'a.Close

'2. Открыть "установка и удаление программ"

'CreateObject("WScript.Shell").Run "control.exe appwiz.cpl"


'3. Создать сценарий обеспечивающий создание в реестре в разделе HKEY_CURRENT_USER собственного раздела

'set WSHShell = WScript.CreateObject("WScript.Shell")

'RegWrite - записываает в реестр заданный параметр или раздел.


'WSHShell.RegWrite "HKCU\NewKey\NewValue", 1, "REG_DWORD"
'WSHShell.Run "regedit",3

'4. Сценарий вывода дата/время

'MsgBox date&vbNewLine&Time

'5. Создать сценарий обеспечивающий чтение реестра содержимого любого раздеела, параметр и значение параметра

'set WshShell1 = wscript.createobject("Wscript.shell")
'Razdel=WshShell1.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\BITS\")
'WScript.Echo Razdel

'6. Открытие excel-я

'set pun2 = wscript.createobject("wscript.shell")
'WScript.Echo "Excel open"
'pun2.run "EXCEL.EXE",2,true

'7. Вывод на экран содержимого файлов config.sys и autoexec.bat

'Dim WshShell, Path
'Set WshShell = WScript.CreateObject("WScript.Shell")
'On Error Resume Next

'Path = "notepad C:\AUTOEXEC.BAT"
'WshShell.Run Path

'Path = "notepad C:\CONFIG.SYS"
'WshShell.Run Path

'8. Cоздать сценарий обеспечивающий открытие любого текстового файла в режиме блокнот

'Dim WshShell
'Dim txt_name
'Set WshShell = WScript.CreateObject("WScript.Shell")
'WScript.Echo "Запускаем Блокнот"
'txt_name=inputbox("Введите имя текстового файла",,"zz.html")
'WshShell.Run "notepad "&txt_name
'WScript.Sleep 1000

'AppActivate - активизирует указанное окно какого-либо приложения. Возвращает True в случае успеха и False в случае неудачи

'WshShell.AppActivate "notepad"

'9. Создание сценария,обеспечивающего вывод на экран содержимого окна "Экран"
'Set objWShell = CreateObject("WScript.Shell")
'objWShell.Run "control.exe desk.cpl", 1
'Set objWShell = Nothing

'10. Подсчёт слов в word документе
'Dim WA, WD, wdStatisticWords, WshShell
'Set WshShell = WScript.CreateObject("WScript.Shell")
'Set WA=WScript.CreateObject("Word.Application")
'Set WD=WA.Documents.Open("D:\1.docx")
'WA.Visible=true
'WScript.Echo "All words: " & WA.ActiveDocument.ComputeStatistics(wdStatisticWords)

'11. Создание ярлыков Excel and Word.

' =========================== EXCEL ===============================

'Dim WSHShell
'Set WSHShell = WScript.CreateObject("WScript.Shell")


'Dim MyShortcut, MyDesktop, DesktopPath

'DesktopPath = WSHShell.SpecialFolders("Desktop")

'Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\excel.lnk")

'MyShortcut.TargetPath = WSHShell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE")
'MyShortcut.WorkingDirectory = WSHShell.ExpandEnvironmentStrings("%windir%")
'MyShortcut.WindowStyle = 4
'MyShortcut.IconLocation = WSHShell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE, 0")
'MyShortcut.Save

'WScript.Echo "Yarlik create"

' =================================== WORD ==========================

'Dim WSHShell
'Set WSHShell = WScript.CreateObject("WScript.Shell")


'Dim MyShortcut, MyDesktop, DesktopPath

'DesktopPath = WSHShell.SpecialFolders("Desktop")

'Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\Word.lnk")

'MyShortcut.TargetPath = WSHShell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE")
'MyShortcut.WorkingDirectory = WSHShell.ExpandEnvironmentStrings("%windir%")
'MyShortcut.WindowStyle = 4
'MyShortcut.IconLocation = WSHShell.ExpandEnvironmentStrings("C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE, 0")
'MyShortcut.Save

'WScript.Echo "Yarlik create"

'12. Получения значения переменной среды ОС Windows

'Set WshShell = WScript.CreateObject("WScript.Shell")
'v = WshShell.ExpandEnvironmentStrings("%SystemRoot%")
'MsgBox v
'13. Просмотр информации о диске

'Dim oFSO ' Объявление переменных
'Dim oDriveInfo

'Set oFSO = CreateObject("Scripting.FileSystemObject") ' Создаём объект для работы с файловой системой

'Set oDriveInfo = oFSO.GetDrive("D:\") ' Получаем диск D:\

' Выводим полученную информацию из переменной oDriveInfo
'MsgBox("Disc - " & oDriveInfo.DriveLetter & vbCrLf & _
' "Metka Disk - " & oDriveInfo.VolumeName & vbCrLf & _
' "Type Disk - " & GetDriveType(oDriveInfo.DriveType) & vbCrLf & _
' "File system - " & oDriveInfo.FileSystem & vbCrLf & _
' "Obiem disk - " & oDriveInfo.TotalSize & " byte" & vbCrLf & _
' "CBobodHoe MecTo for disk - " & oDriveInfo.AvailableSpace & " byte" & vbCrLf & _
' "CBobodHoe MecTo on disk - " & oDriveInfo.FreeSpace & " byte" & vbCrLf & _
' "Seriyniy number disk - " & oDriveInfo.SerialNumber & vbCrLf)
'Function GetDriveType(nType) ' Функция, преобразующая тип диска из числового представления в удобное человеку
' Dim sDriveType

' Select Case nType
' Case 0
' sDriveType = "Heu3BecTHoe ycTpoucTBo"
' Case 1
' sDriveType = "ycTpoucTBo co cmeHHum HocuTelem"
' Case 2
' sDriveType = "HardWare"
' Case 3
' sDriveType = "Setevou disk"
' Case 4
' sDriveType = "CD-ROM"
' Case 5
' sDriveType = "RAM-disk"
' End Select

' GetDriveType = sDriveType
'End Function

'14. ======================== Cоздание файла ========================

'Dim FSO, MyFile
'Set FSO = CreateObject("Scripting.FileSystemObject")
'Set MyFile = fso.CreateTextFile("d:\testfile.txt", true)

' =========================== Удаление файла ========================

'Dim FSO, file
'Set FSO= CreateObject("Scripting.FileSystemObject")
'Set file= FSO.GetFile("d:\testfile.txt")
'file.delete

' =========================== Копирование файла =====================

'Function CopyFiles(FiletoCopy,DestinationFolder)
' Dim fso
' Dim Filepath,WarFileLocation
' Set fso = CreateObject("Scripting.FileSystemObject")
' If Right(DestinationFolder,1) <>"\"Then
' DestinationFolder=DestinationFolder&"\"
' End If
' fso.CopyFile FiletoCopy,DestinationFolder,True
' FiletoCopy = Split(FiletoCopy,"\")

'End Function

'15. Создать пользователя

'Set objNetwork = CreateObject("WScript.Network")
'strComputer = objNetwork.ComputerName

'Set colAccounts = GetObject("WinNT://" & strComputer & "")
'Set objUser = colAccounts.Create("user", "LocalAdmin")
'objUser.SetPassword "Hello123456789"
'objUser.SetInfo

'Set objGroup = GetObject("WinNT://" & strComputer & "/Администраторы,group")
'Set objUser = GetObject("WinNT://" & strComputer & "/LocalAdmin,user")
'objGroup.Add(objUser.ADsPath)

'Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
'UserFlags = objUser.Get("UserFlags")
'objPasswordExpirationFlag = UserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
'objUser.Put "userFlags", objPasswordExpirationFlag
'objUser.SetInfo
