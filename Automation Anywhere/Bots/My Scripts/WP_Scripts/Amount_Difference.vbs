Dim a
Dim b
dim c
'a=17.19
'b=2.34

a = WScript.Arguments(0)
b = WScript.Arguments(1)

c = a-b
c = (Round(c,2))



WScript.StdOut.WriteLine c

