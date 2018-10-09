<%
'function返回一个值给调用者
sub errorHistoryBack(info)
	response.write "<script>alert('"&info&"');history.back();</script>"
	response.end
end sub

'创建一个弹出窗口并且跳转到指定的页面，然后结束语句
sub sussLoctionHref(info,url) 
	response.write "<script>alert('"&info&"');location.href='"&url&"'</script>"
	response.end
end sub

'退出父窗体,跳转到指定的页面，然后结束语句
sub sussLoctionHrefFrame(info,url) 
	response.write "<script>alert('"&info&"');parent.location.href='"&url&"'</script>"
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
	End Select
End Function

'取得周次
Function GetWeekNo(InputDate) 
	dim pytY,pytNewYear,pytNewYearWeek,pytAllDay,pytBanWeek,NumWeek 
	NumWeek = 0 
	pytY = Year(InputDate) 
	pytNewYear=pytY &"-1-1" 
	pytNewYearWeek = Weekday(pytNewYear) 
	pytAllDay = DateDiff("d",pytNewYear,InputDate) 
	pytBanWeek = 8-pytNewYearWeek 
	if pytBanWeek<7 Then 
		NumWeek = 1 
		pytAllDay = pytAllDay - pytBanWeek 
	end if 
	tempx = pytAllDay/7 
	tempx = -Int(-tempx) 
	NumWeek = NumWeek+tempx 
	GetWeekNo = NumWeek 
end Function
%>