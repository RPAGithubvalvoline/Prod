Excel_Path = Wscript.Arguments(0)

PDF_Path = Wscript.Arguments(1)
Dim i
Dim Name
Dim Excel_Path
Dim PDF_Path
'Excel_Path = "C:\Users\c816330\Downloads\DailyStatusReport-25-3-2020_02.xlsx"
'PDF_Path = "C:\Users\c816330\Downloads\"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False   
Set objWorkbook = objExcel.Workbooks.Open(Excel_Path)
'Msgbox ("Done")
For i = 1 To objWorkbook.Worksheets.Count
Set objWorkSheets = objWorkbook.Worksheets(1)
With objExcel.ActiveSheet.PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
End With
objExcel.Application.DisplayAlerts = true
objExcel.DisplayAlerts = true
objExcel.Application.Visible = true
Name = objWorkbook.Worksheets(i).Name
objWorkbook.Worksheets(i).ExportAsFixedFormat 0, PDF_Path&Name&"Excelcase.pdf" ,0, 1, 0,,,0
'Msgbox (Name)
Next
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
SET objExcel = nothing