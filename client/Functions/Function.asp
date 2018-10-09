﻿<%
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

'---------------------创建随机数
Function GetRandomNum()
	dim arrayid(),i,j,blnre,temp,iloop,eloop  
	redim arrayid(0)
	i=0
	iloop=0
	eloop=0
	blnre=false
	randomize
	do while i<1
		temp=int(rnd*2000000)
		if i=0 then
			arrayid(0)=temp
			i=i+1
			iloop=iloop+1
		else
			for j=0 to i-1
				if arrayid(j)=temp then
					blnre=true
					iloop=iloop+1
					exit for
				else
					iloop=iloop+1
				end if
			next

			if blnre=false then
				arrayid(i)=temp
				i=i+1
			else
				blnre=false
			end if
		end if
		eloop=eloop+iloop
		iloop=0
	loop
	GetRandomNum=join(arrayid)
end function


function UpdateJob()


if  ProductName="RA" or  ProductName="Slim" then
				
				JOB_START_QUANTITY690=clng(C4G)+clng(C4B)
				JOB_START_QUANTITY700=clng(C5G)+clng(C5B)
				sql="update  job   set JOB_START_QUANTITY='"&JOB_START_QUANTITY&"', JOB_GOOD_QUANTITY='"&C5G&"'  where job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'" 
				conn.execute(sql)
		
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY&"',GOOD_QUANTITY='"&T1G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'"
				sql=sql+" and STATION_ID='ST00003856'"
				conn.execute(sql)
				
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY690&"',GOOD_QUANTITY='"&C4G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'"
				sql=sql+" and STATION_ID='ST00003857'"
				conn.execute(sql)
				'  response.End()	
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY700&"',GOOD_QUANTITY='"&C5G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'"
				sql=sql+" and STATION_ID='ST00003858'"
				conn.execute(sql)
			
			else
			
				JOB_START_QUANTITY690=clng(C2G)+clng(C2B)
				JOB_START_QUANTITY700=clng(C3G)+clng(C3B)
				sql="update  job   set JOB_START_QUANTITY='"&JOB_START_QUANTITY&"', JOB_GOOD_QUANTITY='"&C3G&"'  where job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'" 
				
		        conn.execute(sql)
		
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY&"',GOOD_QUANTITY='"&T1G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'"
				sql=sql+" and STATION_ID='ST00003856'"
				conn.execute(sql)
				
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY690&"',GOOD_QUANTITY='"&C2G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"'"
				sql=sql+" and STATION_ID='ST00003857'"
				conn.execute(sql)
				
				sql="update JOB_STATIONS set STATION_START_QUANTITY='"&JOB_START_QUANTITY700&"',GOOD_QUANTITY='"&C3G&"'   where  job_number='"&job_number&"' and SHEET_NUMBER='"&sheetnumber&"' "
				sql=sql+" and STATION_ID='ST00003858'"
				conn.execute(sql)
			
			end if
UpdateJob=sql
end function


%>