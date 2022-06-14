On Error Resume Next
Set objSrcExcel = CreateObject("Excel.Application")
objSrcExcel.Visible = False
objSrcExcel.DisplayAlerts = False
Set objSrcExcel1 = CreateObject("Excel.Application")
objSrcExcel1.Visible = False
objSrcExcel1.DisplayAlerts = False
'msgbox ( WScript.Arguments(0))
'msgbox ( WScript.Arguments(1) )
'msgbox ( WScript.Arguments(2) )

'msgbox("CHECK values before input wb, op wb, c code")
Set inputWorkbook = objSrcExcel.Workbooks.Open(WScript.Arguments(0))
'Set inputWorkbook = objSrcExcel.Workbooks.Open("C:\Users\C816328\GL.xlsx")
'msgbox( inputWorkbook )
Set inputWorksheet = inputWorkbook.Worksheets(1)
Set outputWorkbook = objSrcExcel1.Workbooks.Open(WScript.Arguments(1))
'Set outputWorkbook = objSrcExcel1.Workbooks.Open("C:\Users\C816328\Desktop\2.xlsx")
'msgbox( outputWorkbook )
Set outputWorksheet = outputWorkbook.Worksheets(1)
Dim Totalrows
Dim i
Dim CompanyCode
Dim CompanyVerify
Dim Assignment
Dim a
a = 2
CompanyCode = WScript.Arguments(2)
'CompanyCode = "0312"
'msgbox("input wb, op wb, c code")
'msgbox( inputWorkbook )
'msgbox( outputWorkbook )
'msgbox ( CompanyCode )

Totalrows = inputWorksheet.usedrange.rows.count
For i = 2 to Totalrows
CompanyVerify = inputWorkSheet.Cells(i,17).Value
'msgbox("check data from excel ")
'msgbox(CompanyVerify)
'msgbox("check data ishould be 0271 ")
'msgbox(CompanyCode)
If (CompanyVerify = CompanyCode) Then
Assignment = inputWorkSheet.Cells(i,2).Value
Length = Len(Assignment)
If (Length = 13) Then
FirstDigit = Mid(Assignment,1,1)
If (FirstDigit = 8) Then
'MsgBox (Assignment)
outputWorkSheet.Cells(a,1).Value = a-1
outputWorkSheet.Cells(a,2).Value = inputWorkSheet.Cells(i,2).Value  
outputWorkSheet.Cells(a,3).Value = inputWorkSheet.Cells(i,4).Value    
outputWorkSheet.Cells(a,15).Value = inputWorkSheet.Cells(i,6).Value   
outputWorkSheet.Cells(a,6).Value = inputWorkSheet.Cells(i,10).Value    
outputWorkSheet.Cells(a,14).Value = inputWorkSheet.Cells(i,11).Value   
outputWorkSheet.Cells(a,16).Value = inputWorkSheet.Cells(i,14).Value   
a = a+1
End If
End If
End If
Next
inputWorkbook.Save
inputWorkbook.Close
outputWorkbook.Save
outputWorkbook.Close
objSrcExcel.Application.Quit
objSrcExcel1.Application.Quit

If Err.Number <> 0 Then
 'msgbox("error in script")
  'MsgBox "Error # " & CStr(Err.Number) & " " & Err.Description
  WScript.Quit
End If