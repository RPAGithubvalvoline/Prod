input = Wscript.Arguments(0)
output1 = Wscript.Arguments(1)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False 
Set objWorkbook = objExcel.Workbooks.Open(input)
Set objWorkSheets = objWorkbook.Worksheets(1)
With objExcel.ActiveSheet.PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
End With
objExcel.Application.DisplayAlerts = true
objExcel.DisplayAlerts = true
objExcel.Application.Visible = true

objExcel.ActiveSheet.ExportAsFixedFormat 0, output1 ,0, 1, 0,,,0

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
SET objExcel = nothing