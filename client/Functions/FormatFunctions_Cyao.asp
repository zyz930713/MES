<%
function formatdate(vdate,vformat)
if vdate<>"" then
	vyear=year(vdate)
	vmonth=month(vdate)
	vday=day(vdate)
	vhour=hour(vdate)
	vminute=minute(vdate)
	vsecond=second(vdate)
	
	select case lcase(vformat)
	case "ymmdd"
		formatdate=right(vyear,1)&repeatstring(vmonth,"0",2)&repeatstring(vday,"0",2)
	case "yymmdd"
	formatdate=right(vyear,2)&repeatstring(vmonth,"0",2)&repeatstring(vday,"0",2)
	case "ymmddhh"
		formatdate=right(vyear,1)&repeatstring(vmonth,"0",2)&repeatstring(vday,"0",2)&repeatstring(vhour,"0",2)
	case "yyyymmdd"
		formatdate=vyear&"-"&vmonth&"-"&vday
	case "ddmmyyyy"
		formatdate=vday&"/"&vmonth&"/"&vyear
	case "mmddyyyy"
		formatdate=vmonth&"/"&vday&"/"&vyear
	case "yyyymmdd hhnnss"
		formatdate=vyear&"-"&vmonth&"-"&vday&" "&vhour&":"&vminute&":"&vsecond
	case "ddmmyyyy hhnnss"
		formatdate=vday&"/"&vmonth&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case "mmddyyyy hhnnss"
		formatdate=vmonth&"/"&vday&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case else
		formatdate=vdate
	end select
end if
end function

function GetWeekDay(Weekday,TypeCE)
	if TypeCE = 1 then
		select case Weekday
			case 1
				GetWeekDay="Monday"
			case 2
				GetWeekDay="Tuesday"
			case 3
				GetWeekDay="Wednesday"
			case 4
				GetWeekDay="Thursday"
			case 5
				GetWeekDay="Friday" 
			case 6
				GetWeekDay="Saturday"
			case 7
				GetWeekDay="Sunday"
			case else
				GetWeekDay=Weekday
		end select
	elseif TypeCE = 2 then
		select case Weekday
			case 1
				GetWeekDay="星期一"
			case 2
				GetWeekDay="星期二"
			case 3
				GetWeekDay="星期三"
			case 4
				GetWeekDay="星期四"
			case 5
				GetWeekDay="星期五" 
			case 6
				GetWeekDay="星期六"
			case 7
				GetWeekDay="星期日"
			case else
				GetWeekDay=Weekday
		end select
	else
		GetWeekDay=Weekday
	end if
end function

function GetWeekDayEn(Weekday)
	select case Weekday
		case 1
			GetWeekDayEn="Monday"
		case 2
			GetWeekDayEn="Tuesday"
		case 3
			GetWeekDayEn="Wednesday"
		case 4
			GetWeekDayEn="Thursday"
		case 5
			GetWeekDayEn="Friday" 
		case 6
			GetWeekDayEn="Saturday"
		case 7
			GetWeekDayEn="Sunday"
		case else
			GetWeekDayEn=Weekday
	end select
end function

function GetWeekDayCn(Weekday)
	select case Weekday
		case 1
			GetWeekDayCn="星期一"
		case 2
			GetWeekDayCn="星期二"
		case 3
			GetWeekDayCn="星期三"
		case 4
			GetWeekDayCn="星期四"
		case 5
			GetWeekDayCn="星期五" 
		case 6
			GetWeekDayCn="星期六"
		case 7
			GetWeekDayCn="星期日"
		case else
			GetWeekDayCn=Weekday
	end select  
end function

sub errorHistoryBack(info)'function返回一个值给调用者
	response.write "<script>alert('"&info&"');history.back();</script>"
	response.end
end sub

sub sussLoctionHref(info,url) '创建一个弹出窗口并且跳转到指定的页面，然后结束语句
	response.write "<script>alert('"&info&"');location.href='"&url&"'</script>"
	response.end
end sub

'格式化日期格式	
Function FormatTime(s_Time,n_Flag)
Dim y, m, d, h, mi, s
FormatTime = ""
If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
If len(s) = 1 Then s = "0" & s
 	Select Case n_Flag
	Case 1 
		'yyyy-mm-dd hh:mm:ss
		FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		'yyyy-mm-dd
		FormatTime = y & "-" & m & "-" & d
	Case 3 
		'hh:mm:ss
		FormatTime = h & ":" & mi & ":" & s
	Case 4 
		'yyyy年mm月dd日
		FormatTime = y & "年" & m & "月" & d & "日"
	Case 5 
		'yyyymmdd
		FormatTime = y & m & d
	case 6 
		'yyyymmddhhmmss
		FormatTime= y & m & d & h & mi & s
	Case 7
		'hh:mm
		FormatTime = h & ":" & mi
	Case 8
		'hh
		FormatTime = h
	Case 9
		'星期几
		FormatTime = y & "-" & m & "-" & d
		FormatTime = Weekday(FormatTime,2)
		Select Case FormatTime
			Case 1
				FormatTime = "星期一"
			Case 2
				FormatTime = "星期二"
			Case 3
				FormatTime = "星期三"	
			Case 4
				FormatTime = "星期四"	
			Case 5
				FormatTime = "星期五"	
			Case 6
				FormatTime = "星期六"	
			Case 7
				FormatTime = "星期日"
		End Select
		
	End Select
End Function
%>