str = WScript.Arguments(0)
'str = "PY:0914991121:0914991122:0914991159:0914991097:0914991158:0914991120:0914991125:0914991126:0914991128:0914991160:0914991161:0914"
Set RE = New RegExp
RE.pattern = WScript.Arguments(1)
RE.Global = True
RE.IgnoreCase = True
result = ""
Set allMatches = RE.Execute(str)
for each elem in allMatches
result =  elem & result
Next
WScript.StdOut.WriteLine result