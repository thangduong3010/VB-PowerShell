Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, MyFile, FileName, content, arrLines, strLine, insLine, wCnt
Set fso = CreateObject("Scripting.FileSystemObject")

' Open the file for output.
FileName = "F:\Github\VB-PowerShell\Scripts\case.txt"

' Open the file for input.
Set MyFile = fso.OpenTextFile(FileName, ForReading)

content = MyFile.ReadAll()
MyFile.Close
arrLines = Split(content, vbCrLf)
MsgBox content

'wCnt = 0
'
'Set MyFile = fso.OpenTextFile(FileName, ForReading)
'do while not MyFile.AtEndOfStream
'    strLine = MyFile.ReadLine()
'    Wscript.echo "Word to be compared: " & strLine
'	Wscript.echo "Starting Count: " & wCnt
'	
'	for each insLine in arrLines
'		Wscript.echo "Comparing with: " & insLine
'		if strLine = insLine then
'			wCnt = wCnt + 1
'			Wscript.echo "Matched!!!!!!!!!!!!!!!!!!!!!!!!!!"
'		end if
'				
'		Wscript.echo "Count inside: " & wCnt
'	next
'	if wCnt > 0 then
'		Wscript.echo strLine & " has appeared: " & wCnt & " times"
'		wCnt = 0
'	end if
'	wCnt = 0
'loop
'MyFile.Close

Function LoadStringFromFile(filename)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, ForReading)
    LoadStringFromFile = f.ReadAll()
    f.Close
End Function

result = LoadStringFromFile(FileName)
result1 = split(result, vbCrLf)
Wscript.echo result1
