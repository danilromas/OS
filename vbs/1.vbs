Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("Ñ:\ter", True)
a.WriteLine("This is a test.")
a.Close