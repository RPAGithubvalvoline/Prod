Dim ol, fso, folderPath, destPath, f, msg, i

Set ol  = CreateObject("Outlook.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

folderPath = Wscript.Arguments(0)

destPath = Wscript.Arguments(1)

For Each f In fso.GetFolder(folderPath).Files
   
    If LCase(fso.GetExtensionName(f)) = "msg" Then
        
        Set msg = ol.CreateItemFromTemplate(f.Path)
        
        If msg.Attachments.Count > 0 Then
           
            For i = 1 To msg.Attachments.Count
               
                If ((LCase(Mid(msg.Attachments(i).FileName, InStrRev(msg.Attachments(i).FileName, ".") + 1 , 3)) <> "jpg")) AND ((LCase(Mid(msg.Attachments(i).FileName, InStrRev(msg.Attachments(i).FileName, ".") + 1 , 3)) <> "png")) AND ((LCase(Mid(msg.Attachments(i).FileName, InStrRev(msg.Attachments(i).FileName, ".") + 1 , 3)) <> "xml")) then
                     
                    'WScript.Echo f.Name &" -> "& msg.Attachments(i).FileName
                    
                    msg.Attachments(i).SaveAsFile destPath &"\"& msg.Attachments(i).FileName
                End if
          
            Next
        End If
    End If
Next