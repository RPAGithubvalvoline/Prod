dim output
dim fromDate
dim toDate
fromDate = CDate("2020/03/10")
toDate = CDate("2020/03/26")
'response.write(DateDiff("w",fromDate,toDate,vbMonday))
output = DateDiff("d",fromDate,toDate)
msgbox(output )