Set objSrcExcel = CreateObject("Excel.Application")
objSrcExcel.Visible = False
objSrcExcel.DisplayAlerts = False
Set inputWorkbook = objSrcExcel.Workbooks.Open(WScript.Arguments(0))
'Set inputWorkbook = objSrcExcel.Workbooks.Open("C:\Users\choudharym\Downloads\GL.xlsx")
Set inputWorksheet = inputWorkbook.Worksheets(1)
Set outputWorkbook = objSrcExcel.Workbooks.Open(WScript.Arguments(1))
'Set outputWorkbook = objSrcExcel.Workbooks.Open("C:\Users\choudharym\Desktop\New folder\1.xlsx")
Set outputWorksheet = outputWorkbook.Worksheets(1)
Dim Totalrows
Dim i
Dim CompanyCode
Dim CompanyVerify
Dim Assignment
Dim a
a = 1
'CompanyCode = "0271"
CompanyCode = WScript.Arguments(2)
Totalrows = inputWorksheet.usedrange.rows.count
For i = 2 to Totalrows
CompanyVerify = inputWorkSheet.Cells(i,17).Value
If (CompanyVerify = CompanyCode) Then
Assignment = inputWorkSheet.Cells(i,2).Value
Length = Len(Assignment)
If (Length = 13) Then
FirstDigit = Mid(Assignment,1,1)
If (FirstDigit = 8) Then
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