On Error Resume Next
Const ForReading = 1

Const ForWriting = 2

Dim StrLength1

Dim StrLength2
Dim StrLength3
Dim DataLine1
Dim DataLine2
Dim DataLine3
Dim ConcatData
Dim vDataMatched 
vDataMatched = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForReading)


strSearchString = objFile.ReadAll

objFile.Close


Set objRegEx = CreateObject("VBScript.RegExp")


objRegEx.Global = True



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

  DataLine3 = DataLine2
  DataLine2 = DataLine1 
  DataLine1   = strLine

 'msgbox ("data3")
  'msgbox ( DataLine3 )
  

StrLength2 = StrLength1
StrLength1 = Len(strLine)


   'If StrLength1 = "" Then
        
    '    StrLength2 = ""
   
   ' else
    '     StrLength2 = Len(strLine) 
     '    StrLength1 = ""

  'End If
 
  StrLength3 = StrLength1 + StrLength2

 'msgbox ("string length 3")
  'msgbox (StrLength3)

    If (StrLength3) = 9 Then
      'msgbox ( " Condition matched ConcatData will be " )
     ConcatData = DataLine2   & DataLine1   
      'msgbox ( ConcatData )
        vDataMatched = 1           

        strNewContents = strNewContents & ConcatData  & vbCrLf
        StrLength2 = 0
        StrLength1 = 0
       
   if DataLine3 <> "" Then
           strNewContents = strNewContents & DataLine3  & vbCrLf
     End If
      DataLine1   = "" 
      DataLine2   = ""
      DataLine3   = ""      
   else
      vDataMatched = 0
    'msgbox("data line 3 concat data will be" )
    'msgbox( DataLine3 )
      if DataLine3 <> "" Then
           strNewContents = strNewContents & DataLine3  & vbCrLf
     End If

   End If
Loop

 if vDataMatched = 0  Then
   if DataLine2 <> "" Then
             strNewContents = strNewContents & DataLine2  & vbCrLf
    End If

   if DataLine1 <> "" Then
           strNewContents = strNewContents & DataLine1  & vbCrLf
  End If
 End If


'msgbox ("string length 1")
'msgbox (StrLength1)
'msgbox ("string length 2")
'msgbox (StrLength2) 
'msgbox ("string length 3")
'msgbox (StrLength3)

'msgbox ("data line  1")
'msgbox ( DataLine1)
'msgbox ("data line  2")
'msgbox ( DataLine2)
'msgbox ("data line  3")
'msgbox ( DataLine3)
 


'msGbox("NEW CONTENTS")
'msGBox(strNewContents)

Set objFile = objFSO.OpenTextFile(WScript.Arguments(0), ForWriting)

objFile.Write strNewContents


objFile.Close
'END

