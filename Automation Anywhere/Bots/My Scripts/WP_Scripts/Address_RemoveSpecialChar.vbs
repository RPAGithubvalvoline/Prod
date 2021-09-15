on Error Resume Next
dim newvalue

inputString = WScript.Arguments(0)

'msgbox strSearchString 

Set objRegEx = CreateObject("VBScript.RegExp")


objRegEx.Global = True

objRegEx.Pattern = "[^A-Za-z]"




newvalue = objRegEx.Replace(inputString,chr(32))


'msgbox newvalue 

WScript.StdOut.WriteLine newvalue



