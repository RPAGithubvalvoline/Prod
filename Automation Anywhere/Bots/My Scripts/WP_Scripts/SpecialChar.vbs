Const ForReading = 1

Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForReading)


strSearchString = objFile.ReadAll

objFile.Close


Set objRegEx = CreateObject("VBScript.RegExp")


objRegEx.Global = True

'objRegEx.Pattern = "[^A-Za-z0-9\n\r]"

objRegEx.Pattern = "[^0-9\n\r]"


strSearchString = objRegEx.Replace(strSearchString,chr(13)+chr(10))


Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForWriting)

objFile.WriteLine strSearchString
objFile.Close

'START

Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForReading)




Do Until objFile.AtEndOfStream

    strLine = objFile.Readline

    strLine = Trim(strLine)

    If Len(strLine) > 0 Then

        strNewContents = strNewContents & strLine & vbCrLf

    End If

Loop

objFile.Close

Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForWriting)

objFile.Write strNewContents


objFile.Close
'END

