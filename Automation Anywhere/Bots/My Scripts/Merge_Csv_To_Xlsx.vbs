tgtxlsfile = Wscript.Arguments(0)
srccsvfile = Wscript.Arguments(1)

Set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False
objExcel.DisplayAlerts = False
objExcel.Application.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(tgtxlsfile)
Set objWorkbook1 = objExcel.Workbooks.Open(srccsvfile)

Set objIFAP = objWorkbook1.Worksheets(1)
vBlank = objIFAP.usedrange.rows.count
objIFAP.Range("A1:M"&vBlank&"").Copy

Set xlApp = CreateObject("Excel.Application")

xlApp.Application.DisplayAlerts = False
xlApp.DisplayAlerts = False
xlApp.Application.Visible = False

xlApp.Workbooks.Open(tgtxlsfile)
Set objFS = objWorkbook.ActiveSheet
'Add Sheet To Last
Set objFS = objWorkbook.Worksheets.Add(, objWorkbook.Worksheets(objWorkbook.Worksheets.Count))
'Rename it
objFS.Name = "Scanning_Daily_Status_Report"
objFS.Range("A1:M"&vBlank&"").PasteSpecial
objFS.Range("A1:N1").Font.Bold = True
objWorkbook.SaveAs tgtxlsfile 
objWorkbook.Close
objWorkbook1.Close
objExcel.Quit
xlApp.Quit
set objExcel = nothing
Set xlApp = nothing