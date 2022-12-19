Dim FSO, MyFile
Set FSO = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile("d:\testfile.txt", true)

\\delete

Dim FSO, file
Set FSO= CreateObject("Scripting.FileSystemObject")
Set file= FSO.GetFile("d:\testfile.txt")
file.delete

\\copy

Function CopyFiles(FiletoCopy,DestinationFolder)
Dim fso
Dim Filepath,WarFileLocation
Set fso = CreateObject("Scripting.FileSystemObject")
If Right(DestinationFolder,1) <>"\"Then
DestinationFolder=DestinationFolder&"\"
End If
fso.CopyFile FiletoCopy,DestinationFolder,True
FiletoCopy = Split(FiletoCopy,"\")

End Function