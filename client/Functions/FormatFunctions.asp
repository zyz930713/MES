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

function getmin(a,b)
	if a<b then
	getmin=a
	else
	getmin=b
	end if
end function

function repeatstring(str,ex,num)
	repeatstring=string(num-len(str),ex)&str
end function


function getDays(Weekdays)

		select case Weekdays
		  
		  case "monday" 
		  getDays="1"
		  case "tuesday" 
		  getDays="2"
		  case "wednesday"  
		  getDays="3"
		  case "thursday"  
		  getDays="4"
		  case "friday"  
		  getDays="5"
		  case "saturday"  
		  getDays="6"
		  case "sunday" 
		  getDays="7"
		end  select  
end function


function getprodShift(prodShift)
	select case prodShift
	case "A"
		getprodShift="1"
	case "B"
		getprodShift="2"
	case "C"
		getprodShift="3"
	case "D"
		getprodShift="4"					
	end select
end function




function getOneCharMon(strMon)
	select case strMon
	case "01"
		getOneCharMon="1"
	case "02"
		getOneCharMon="2"
	case "03"
		getOneCharMon="3"
	case "04"
		getOneCharMon="4"
	case "05"
		getOneCharMon="5"
	case "06"
		getOneCharMon="6"
	case "07"
		getOneCharMon="7"
	case "08"
		getOneCharMon="8"
	case "09"
		getOneCharMon="9"
	case "10"
		getOneCharMon="A"
	case "11"
		getOneCharMon="B"
	case "12"
		getOneCharMon="C"					
	end select
end function

function N10toC62(b, bt) 
    If bt<2 Or bt>62 Then
        bt = 16 '默认为 16 进制
    End If
	'2进制 0-1
    '8进制 0-7          可以用 OCT 函数代替
    '16进制 0-9 A-F     可以用 HEX 函数代替
    '36进制 0-9 A-Z
    '62进制 0-9 A-Z a-z
    '都不对，就用16进制，如果输入数据不符合要求，则出错    
    Dim a 
    Dim a1 
    Dim s
    Do while b<>0
		a= b mod bt
		if a>=0 and a<=9 then
			a1 = CStr(a)
		end if

		if a>=10 and a<=35 then
			a1 = Chr(a + 55)
		end if
	
		if a>=36 and a<=61 then
			a1 = Chr(a + 61)
		end if

'        Select Case a
'        Case 0 To 9
'            a1 = CStr(a)
'        Case 10 To 35
'            a1 = Chr(a + 55)
'        Case 36 To 61
'            a1 = Chr(a + 61)
'        End Select
        s = a1 & s
        b = Int(b / bt)
    Loop 
    N10toC62 = s
end Function

function formatlongstring(str,splitchar,num)
	mystr=""
	if  str<>"" then
		fj=1
		for fi=1 to len(str)
			if fj=num then
			fj=1
			mystr=mystr&mid(str,fi,1)&splitchar
			else
			fj=fj+1
			mystr=mystr&mid(str,fi,1)
			end if
		next
	end if
	formatlongstring=mystr
end function

function checkJob(jobnumber,sheetnumber,jobtype,repeated_sequence)
'set rsch=server.CreateObject("adodb.recordset")
'	SQLch="select STATIONS_INDEX from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
'	 
'	rsch.open SQLch,conn,1,3
'	if not rsch.eof then
'		stations_index=rsch("STATIONS_INDEX")
'	end if
'	rsch.close
'
'a_stations_index=split(stations_index,",")
'current_station=0
'total_defect_quantity=0
'for i=0 to ubound(a_stations_index)
'		if i=0 then
'			this_defect_quantity=0
'			SQLch="select ACTION_VALUE from JOB_ACTIONS a , action b where a.action_id=b.nid(+) and a.JOB_NUMBER='"&jobnumber&"' and a.SHEET_NUMBER="&sheetnumber&" and a.JOB_TYPE='"&jobtype&"' and a.STATION_ID='"&a_stations_index(i)&"' and a.REPEATED_SEQUENCE="&repeated_sequence&" and b.mother_action_id='AN00000061'"
'			rsch.open SQLch,conn,1,3
'			if not rsch.eof then
'			first_action_start_quantity=clng(rsch("ACTION_VALUE"))
'			end if
'			rsch.close
'			SQLch="select STATION_START_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&repeated_sequence
'			 
'			
'			rsch.open SQLch,conn,1,3
'			if not rsch.eof then
'				first_station_start_quantity=clng(rsch("STATION_START_QUANTITY"))
'				if first_station_start_quantity<>first_action_start_quantity then
'				error_string=error_string&"<br>[Start Number Error] Job Station/First Action:"&clng(rsch("STATION_START_QUANTITY"))&"/"&clng(first_action_start_quantity)
'				first_station_start_quantity=first_action_start_quantity
'				end if
'				next_start_quantity=first_station_start_quantity
'			else
'			error_string=error_string&"<br>[No Start Number Error]"
'			rsch.close
'			exit for
'			end if
'			rsch.close
'		end if
'		SQLch="select nvl(sum(DEFECT_QUANTITY),0) as DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE='"&repeated_sequence&"'"
'		rsch.open SQLch,conn,1,3
'		this_defect_quantity=cint(rsch("DEFECT_QUANTITY"))
'		rsch.close
'		this_start_quantity=next_start_quantity
'		this_good_quantity=this_start_quantity-this_defect_quantity
'		total_defect_quantity=total_defect_quantity+this_defect_quantity
'		next_start_quantity=this_good_quantity
'	 
'next
'
'SQLch="select * from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
'session("aerror")=SQL
' 
'rsch.open SQLch,conn,1,3
'if not rsch.eof then
'old_total_defect_quantity=cint(rsch("JOB_DEFECTCODE_QUANTITY"))
'diff=total_defect_quantity-old_total_defect_quantity
'	if diff<>0 then
'		set JMail=server.CreateObject("JMail.Message") 
'		JMail.ContentType = "text/html"
'		JMail.Charset ="gb2312"
'		JMail.From = "B_E@knowles.com"
'		'JMail.AddRecipient "dickens.xu@knowles.com"
'		JMail.AddRecipient "jack.zhang@knowles.com"
'		'JMail.AddRecipient "howard.zhang@knowles.com"&" on "&request.ServerVariables("HTTP_HOST")
'		JMail.Subject = "Station Error for "&jobnumber&"-"&repeatstring(sheetnumber,"0",3)&" "&rsch("PART_NUMBER_TAG")&" on "&request.ServerVariables("HTTP_HOST")
'		JMail.Body = "<a href=""http://KEB-barcode:8081/Job/SubJobs/JobDetail.asp?jobnumber="&jobnumber&"&sheetnumber="&sheetnumber&"&jobtype=N"">"&jobnumber&"-"&repeatstring(sheetnumber,"0",3)&"</a><br>"&error_string&"<br>old Start:"&rsch("JOB_START_QUANTITY")&" | old good:"&rsch("JOB_GOOD_QUANTITY")&" | old bad:"&rsch("JOB_DEFECTCODE_QUANTITY")&"<br> new start:"&first_station_start_quantity&" | new good:"&first_station_start_quantity-total_defect_quantity&" | new bad:"&total_defect_quantity & "<br>SQLch:" & SQLch
'		JMail.Send (application("MailServer"))
'		set JMail = Nothing
'	end if
'end if
'rsch.close
'
'set rsch=nothing
end function

function getFileTypeIcon(files_name)
	if files_name<>"" then 
	file_string=split(files_name,".")
	file_ext=lcase(file_string(ubound(file_string)))
	end if	
	if instr("jpg,gif,png",file_ext)>0 then
	getFileTypeIcon="FILE_IMAGE.jpg"
	elseif instr("doc,docx",file_ext)>0 then
	getFileTypeIcon="FILE_WORD.jpg"
	elseif instr("xls,xlsx",file_ext)>0 then
	getFileTypeIcon="FILE_EXCEL.jpg"
	elseif instr("pps,ppt,pptx",file_ext)>0 then
	getFileTypeIcon="FILE_POWERPOINT.jpg"
	elseif instr("vsd",file_ext)>0 then
	getFileTypeIcon="FILE_VISIO.jpg"
	elseif instr("pdf",file_ext)>0 then
	getFileTypeIcon="FILE_PDF.jpg"
	elseif instr("zip",file_ext)>0 then
	getFileTypeIcon="FILE_ZIP.jpg"	
	elseif instr("rar",file_ext)>0 then
	getFileTypeIcon="FILE_RAR.jpg"	
	elseif instr("mp3,wmv,wma",file_ext)>0 then
	getFileTypeIcon="FILE_MSM.jpg"
	elseif instr("swf",file_ext)>0 then
	getFileTypeIcon="FILE_SWF.jpg"		
	end if
end function
%>