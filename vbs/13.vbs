Dim oFSO
Dim oDriveInfo

Set oFSO = CreateObject("Scripting.FileSystemObject") ' Создаём объект для работы с файловой системой

Set oDriveInfo = oFSO.GetDrive("D:\") ' Получаем диск D:\


MsgBox("Disc - " & oDriveInfo.DriveLetter & vbCrLf & _
"Metka Disk - " & oDriveInfo.VolumeName & vbCrLf & _
"Type Disk - " & GetDriveType(oDriveInfo.DriveType) & vbCrLf & _
"File system - " & oDriveInfo.FileSystem & vbCrLf & _
"Obiem disk - " & oDriveInfo.TotalSize & " byte" & vbCrLf & _
"CBobodHoe MecTo for disk - " & oDriveInfo.AvailableSpace & " byte" & vbCrLf & _
"CBobodHoe MecTo on disk - " & oDriveInfo.FreeSpace & " byte" & vbCrLf & _
"Seriyniy number disk - " & oDriveInfo.SerialNumber & vbCrLf)
Function GetDriveType(nType) ' Функция, преобразующая тип диска из числового представления в удобное человеку
Dim sDriveType

Select Case nType
 Case 0
sDriveType = "Heu3BecTHoe ycTpoucTBo"
 Case 1
 sDriveType = "ycTpoucTBo co cmeHHum HocuTelem"
 Case 2
 sDriveType = "HardWare"
 Case 3
 sDriveType = "Setevou disk"
 Case 4
 sDriveType = "CD-ROM"
 Case 5
sDriveType = "RAM-disk"
 End Select

GetDriveType = sDriveType
End Function
