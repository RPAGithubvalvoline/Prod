Dim rx ,matches, match,line, re
Const ForReading = 1

'vFile = "C:\Users\sarmar\Desktop\Horesh\Text\3.txt"
'vOutputFile = "C:\Users\sarmar\Desktop\Horesh\Text\Out.txt"

vFile = WScript.Arguments(0)
vOutputFile = WScript.Arguments(1)
vKeyValue = WScript.Arguments(2)
vKeyWord = WScript.Arguments(3)

MsgBox vKeyWord

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If objFSO.FileExists(vOutputFile)  Then
	objFSO.DeleteFile vOutputFile
End If


Set file = objFSO.OpenTextFile(vFile,ForReading)  
Set file2 = objFSO.OpenTextFile(vFile,ForReading)  

Set rx = New RegExp
rx.Global= True
rx.IgnoreCase = True
rx.Global     = False
rx.Pattern= ""&vKeyValue&"+\d{6}"   
'MsgBox rx.Pattern


Set re1 = New RegExp
re1.IgnoreCase = True
re1.Global     = False
re1.Pattern = "[ ]{2,}"

Set re = New RegExp
re.IgnoreCase = True
re.Global     = False
re.Pattern    = ""&vKeyWord&""


MsgBox re.Pattern

Set objFileToWrite = objFSO.OpenTextFile(vOutputFile,2,true)

vInnerCounter = 0	
Do Until file.AtEndOfStream
    line = file.ReadLine
	Flag = re.Test(line)
	If line <> "" And Flag = "True"  Then
		flag2 = vInnerCounter
		Exit Do
	End If
	vInnerCounter = vInnerCounter + 1
Loop

vInnerCounter = 0	
Do Until file2.AtEndOfStream
    line = file2.ReadLine
	If line <> "" And vInnerCounter >= flag2  Then
		Set matches = rx.Execute(line)
		For Each match In matches
			results = Trim(re1.Replace(line, "|"))
			'MsgBox results
			arr = Split(results,"|")			
			
			'text = Arr(0) &" | "&  Arr(1)
			objFileToWrite.WriteLine(results)					
		Next		

	End If
	vInnerCounter = vInnerCounter + 1
Loop






