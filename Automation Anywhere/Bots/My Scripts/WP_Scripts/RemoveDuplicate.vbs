
vInputPath = Wscript.Arguments.Item(0)

vOutputPath = Wscript.Arguments.Item(1)

Const ForReading = 1
Const ForWriting = 2

Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile(vInputPath, ForReading)

Do Until objFile.AtEndOfStream
strName = objFile.ReadLine
If Not objDictionary.Exists(strName) Then
objDictionary.Add strName, strName
End If
Loop
objFile.Close

Set MyFile = objFSO.CreateTextFile(vOutputPath, True)
For Each strKey in objDictionary.Keys
MyFile.Write strKey & vbCrLf
Next