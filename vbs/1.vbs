Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("�:\ter", True)
a.WriteLine("This is a test.")
a.Close