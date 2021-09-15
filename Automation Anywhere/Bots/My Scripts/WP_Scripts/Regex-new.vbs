Dim rx ,matches, match,line, re
Const ForReading = 1

vFile = "C:\Users\sarmar\Desktop\Horesh\Text\3.txt"
vOutputFile = "C:\Users\sarmar\Desktop\Horesh\Text\Out.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If objFSO.FileExists(vOutputFile)  Then
	objFSO.DeleteFile vOutputFile
End If

'Set oTxtFile = objFSO.CreateTextFile(vOutputFile)

Set file = objFSO.OpenTextFile(vFile,ForReading)  
Set file2 = objFSO.OpenTextFile(vFile,ForReading)  

Set rx = New RegExp
rx.Global= True
rx.IgnoreCase = True
rx.Global     = False
rx.Pattern= "914+\d{6}"   


'Set re = New RegExp
're.IgnoreCase = True
're.Global     = False
're.Pattern = "[ ]{2,}"

Set re = New RegExp
re.Pattern    = "Beløb"
re.IgnoreCase = True
re.Global     = False



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
			results = Trim(re.Replace(line, "|"))
			MsgBox results
			arr = Split(results,"|")			
			
			'text = Arr(0) &" | "&  Arr(1)
			objFileToWrite.WriteLine(results)					
		Next		

	End If
	vInnerCounter = vInnerCounter + 1
Loop

		
MsgBox("done")







