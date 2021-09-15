On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
'Filepath="C:\Users\c816330\Desktop\ProcessReportFolder_NAm\Report Folder\DailyStatusReport-11-2-2020_01 - Copy.csv"
'DestinationFilepath ="C:\Users\c816330\Desktop\ProcessReportFolder_NAm\Report Folder\DailyStatusReport-11-2-2020_01 - Copy"
Filepath = Wscript.Arguments(0)
DestinationFilepath = Wscript.Arguments(1)


If Not fso.FileExists(Filepath) Then

                WScript.StdOut.WriteLine("File not Found")

                WScript.Quit 0

End If


Set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False
objExcel.DisplayAlerts = False
objExcel.Application.Visible = False

'objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open(Filepath)

x = 2

  Do Until objExcel.Cells(x,1).Value = ""
     
    if InStr(1, objExcel.Cells(x,5).Value, "Exception") > 0 then
    'if objExcel.Cells(x,6).Value <>"None" then 
      'msgbox(objExcel.Cells(x,6).Value)
        vRange = "A" &x&":I"&x
        objExcel.Range(vRange).Interior.ColorIndex = 5.5

    end if

   x = x + 1

  Loop

objWorkbook.sheets(1).Range("A1:I1").Font.Bold = True

objWorkbook.saveas(DestinationFilepath),51

objWorkbook.close

objExcel.quit

set objExcel = nothing

if err.number <> 0 then

                WScript.StdOut.WriteLine(err.Description)

end if
WScript.Quit