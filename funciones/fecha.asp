<%public function fecha_mysql(vfecha)
	afecha = split(vfecha,"-")
	for i=ubound(afecha) to 0 step -1
		nfecha = nfecha +"-"+ afecha(i)
	next
	fecha_mysql =mid(nfecha,2,len(nfecha)-1)
end function

public function fecha_mysql2(vfecha,separador)
	afecha = split(vfecha,separador)
	for i=ubound(afecha) to 0 step -1
		nfecha = nfecha +"-"+ afecha(i)
	next
	fecha_mysql2 =mid(nfecha,2,len(nfecha)-1)
end function

public function fecha_mysql3(vfecha)
	afecha = split(vfecha,"-")
	for i=ubound(afecha) to 0 step -1
		nfecha = nfecha & afecha(i)
	next
	fecha_mysql3 =mid(nfecha,1,len(nfecha))
end function

public function fecha_mysql33(vfecha)
	afecha = split(vfecha,"-")
	for i=ubound(afecha) to 0 step -1
		nfecha = afecha(i) & "-" & nfecha 
	next
	fecha_mysql33 =mid(nfecha,1,len(nfecha)-10)
end function

public function fecha_mysql4(vfecha)
	afecha = split(vfecha,"/")
	for i=ubound(afecha) to 0 step -1
		nfecha = nfecha & afecha(i)
	next
	fecha_mysql4 =mid(nfecha,1,len(nfecha))
end function

function formatDate(format, intTimeStamp)
dim unUDate, A

' Test to see if intTimeStamp looks valid. If not, they have passed a normal date
if not (isnumeric(intTimeStamp)) then
if isdate(intTimeStamp) then
intTimeStamp = DateDiff("S", "01/01/1970 00:00:00", intTimeStamp)
else
response.write "Date Invalid"
exit function
end if
end if

if (intTimeStamp=0) then
unUDate = now()
else
unUDate = DateAdd("s", intTimeStamp, "01/01/1970 00:00:00")
end if

unUDate = trim(unUDate)

dim startM : startM = InStr(1, unUDate, "/", vbTextCompare) + 1
dim startY : startY = InStr(startM, unUDate, "/", vbTextCompare) + 1
dim startHour : startHour = InStr(startY, unUDate, " ", vbTextCompare) + 1
dim startMin : startMin = InStr(startHour, unUDate, ":", vbTextCompare) + 1

dim dateDay : dateDay = mid(unUDate, 1, 2)
dim dateMonth : dateMonth = mid(unUDate, startM, 2)
dim dateYear : dateYear = mid(unUDate, startY, 4)
dim dateHour : dateHour = mid(unUDate, startHour, 2)
dim dateMinute : dateMinute = mid(unUDate, startMin, 2)
dim dateSecond : dateSecond = mid(unUDate, InStr(startMin, unUDate, ":", vbTextCompare) + 1, 2)

format = replace(format, "%Y", right(dateYear, 4))
format = replace(format, "%y", right(dateYear, 2))
format = replace(format, "%m", dateMonth)
format = replace(format, "%n", cint(dateMonth))
format = replace(format, "%F", monthname(cint(dateMonth)))
format = replace(format, "%M", left(monthname(cint(dateMonth)), 3))
format = replace(format, "%d", dateDay)
format = replace(format, "%j", cint(dateDay))
format = replace(format, "%h", mid(unUDate, startHour, 2))
format = replace(format, "%g", cint(mid(unUDate, startHour, 2)))

if (cint(dateHour) > 12) then
A = "PM"
else
A = "AM"
end if
format = replace(format, "%A", A)
format = replace(format, "%a", lcase(A))

if (A = "PM") then format = replace(format, "%H", left("0" & dateHour - 12, 2))
format = replace(format, "%H", dateHour)
if (A = "PM") then format = replace(format, "%G", left("0" & cint(dateHour) - 12, 2))
format = replace(format, "%G", cint(dateHour))

format = replace(format, "%i", dateMinute)
format = replace(format, "%I", cint(dateMinute))
format = replace(format, "%s", dateSecond)
format = replace(format, "%S", cint(dateSecond))
format = replace(format, "%L", WeekDay(unUDate))
format = replace(format, "%D", left(WeekDayName(WeekDay(unUDate)), 3))
format = replace(format, "%l", WeekDayName(WeekDay(unUDate)))
format = replace(format, "%U", intTimeStamp)
format = replace(format, "11%O", "11th")
format = replace(format, "1%O", "1st")
format = replace(format, "12%O", "12th")
format = replace(format, "2%O", "2nd")
format = replace(format, "13%O", "13th")
format = replace(format, "3%O", "3rd")
format = replace(format, "%O", "th")

formatDate = format

end function%>