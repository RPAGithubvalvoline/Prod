Dim rx ,matches, match,line, re
Const ForReading = 1
Const ForWriting = 2

'vFile = "C:\Users\C816328\Desktop\Amount_Samples\Betaalspecificatie.txt"
'vOutputFile = "C:\Users\C816328\Desktop\Amount_Samples\Betaalspecificatie_Amount_output.txt"
'vKeyValue =  "914" 
'vKeyWord = "Bedrag"
'vInvoicePos = "0"

vFile = WScript.Arguments(0)
vOutputFile = WScript.Arguments(1)
vKeyValue = WScript.Arguments(2)
vKeyWord = WScript.Arguments(3)
vInvoicePos = WScript.Arguments(4)

Msgbox vKeyWord
Msgbox vInvoicePos


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

Set re1 = New RegExp
re1.IgnoreCase = True
re1.Global     = True
re1.Pattern = "[ ]{2,}"

Set re = New RegExp
re.IgnoreCase = False
re.Global     = False
re.Pattern    =  ""&vKeyWord&""

Set objFileToWrite = objFSO.OpenTextFile(vOutputFile,ForWriting,true)
vInnerCounter = 0	
Do Until file.AtEndOfStream
    line = file.ReadLine
	Flag = re.Test(line)
	If line <> "" And Flag = "True"  Then	
		vHeader = Trim(re1.Replace(line,"|"))
		vHeaderArray = Split(vHeader,"|")	
		For i=lbound(vHeaderArray) to ubound(vHeaderArray)			
			value1 = vHeaderArray(i) 			
			If value1 = vKeyWord  Then
				vHeaderPos1 = i 
				exit for
			End If
		Next		
		flag2 = vInnerCounter
		Exit Do
	End If
	vInnerCounter = vInnerCounter + 1
Loop
vInnerCounter = 0	
Do Until file2.AtEndOfStream
    line = file2.ReadLine	
	If line <> "" And vInnerCounter >= flag2  Then
		test2 = rx.Test(line)		
		If line <> "" And test2 = "True"  Then
			results = Trim(re1.Replace(line,"|"))
			rr = Split(results,"|")				
			vFinalText = rr(vInvoicePos) & " | " & rr(vHeaderPos1)					
			objFileToWrite.WriteLine(vFinalText)			
		End If
	End If
	vInnerCounter = vInnerCounter + 1
Loop
MsgBox("done")