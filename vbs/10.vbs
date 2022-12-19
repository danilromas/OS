Dim WWORD, word, KolSlov, shell
Set shell = WScript.CreateObject("WScript.Shell")
Set WWORD=WScript.CreateObject("Word.Application")
Set word=WWORD.Documents.Open("D:\1.docx")
WWORD.Visible=true
WScript.Echo "All words: " & WWORD.ActiveDocument.ComputeStatistics(KolSlov)