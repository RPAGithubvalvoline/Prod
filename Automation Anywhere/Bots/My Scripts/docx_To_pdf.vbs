Const wdExportAllDocument = 0
Const wdExportOptimizeForPrint = 0
Const wdExportDocumentContent = 0
Const wdExportFormatPDF = 17
Const wdExportCreateHeadingBookmarks = 1

if  Wscript.Arguments.Count > 0 Then
    ' Get the running instance of MS Word. If Word is not running, Create it
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set objWord = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(WScript.Arguments(0))
    Set objDoc = objWord.Documents.Open(WScript.Arguments(0),,TRUE)

    'Export to PDF using preferred settings
    pdf = objWord.ActiveDocument.ExportAsFixedFormat( _
        WScript.Arguments(1), _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, _
        wdExportAllDocument,,, _
        wdExportDocumentContent, _
        False, True, _
        wdExportCreateHeadingBookmarks _
    )

    'Quit MS Word
    objWord.DisplayAlerts = False
    objWord.Quit(False)
    set objWord = nothing
    set objFSO = nothing
Else
    msgbox("You must select a file to convert")
End If
