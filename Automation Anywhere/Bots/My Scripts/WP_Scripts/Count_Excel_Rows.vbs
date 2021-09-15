on error resume next
Dim lngNoOfRows 
Dim excel_path 

excel_path = (WScript.Arguments(0))

'Path= "C:\Users\C816328\Desktop\RAJIB\output\123.xlsx"

Set objexl  = Createobject("Excel.application")

objexl.visible = False

set objwkb = objexl.workbooks.open(path)

set objsht = objwkb.sheets(1)

'msgbox objsht.usedrange.rows.count  

lngNoOfRows = objsht.Range("A65536").End(-4162).Row

msgbox lngNoOfRows 

WScript.StdOut.WriteLine lngNoOfRows

objexl.Application.Quit