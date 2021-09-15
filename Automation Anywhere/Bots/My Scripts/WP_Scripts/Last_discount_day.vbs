
On Error Resume Next

Dim vOutput
'dd_payment = 19
dd_payment  = int(WScript.Arguments(0))

'mm_payment = 12

mm_payment = int(WScript.Arguments(1))

'yy_payment = 2019

yy_payment = int(WScript.Arguments(2))

'dd_bline = 10

dd_bline = int(WScript.Arguments(3))

'mm_bline = 12

mm_bline = int(WScript.Arguments(4))
'yy_bline = 2019

yy_bline= int(WScript.Arguments(5))
'Grace_period = 4
Grace_period = int(WScript.Arguments(6))
'Discount_days = 10
Discount_days = int(WScript.Arguments(7))

'payment_date = dd_payment & "/" & mm_payment & "/" & yy_payment
payment_date = mm_payment & "/" & dd_payment & "/" & yy_payment
'msgbox(payment_date)
payment_date = cDate(payment_date)
'msgbox("AFTER")
'msgbox(payment_date)
'Bline_date = dd_bline & "-" & mm_bline & "-" & yy_bline
Bline_date = mm_bline & "/" & dd_bline & "/" & yy_bline
Bline_date = cDate(Bline_date)

Total_adding_days = Grace_period + Discount_days

Last_payment_date = DateAdd("d",Total_adding_days,Bline_date)

'Msgbox("last pymnet date and total adding days and bline date")
'msgbox(Last_payment_date )
'msgbox(Total_adding_days)
'msgbox(Bline_date )


weekday_last_Payment_date = WeekdayName(Weekday(Last_payment_date))

if weekday_last_Payment_date = "Sunday" then
    Last_payment_date = DateAdd("d",-2,Last_payment_date)

	
else if weekday_last_Payment_date = "Saturday" then
    Last_payment_date = DateAdd("d",-1,Last_payment_date)
End if
End if
'msgbox("last payment date should be gretate than and Payment Date")
'msgbox(Last_payment_date)
'msgbox(payment_date)

payment_date = cDate(payment_date)
'Last_payment_date=cDate(Last_payment_date)

'msgbox("last payment date should be gretate than and Payment Date")

'msgbox(Last_payment_date)
'msgbox(payment_date)

if Last_payment_date >= payment_date then

 vOutput = "Y"
 WScript.StdOut.WriteLine (vOutput)
 'msgbox vOutput
else
 vOutput = "N"
 WScript.StdOut.WriteLine (vOutput)
 'msgbox vOutput
End if

