curMonth=right("00" & month(now),2)
mnthname = left(MonthName(curMonth),3)
Args2=curMonth & "-" & mnthname
Wscript.Stdout.writeline Args2