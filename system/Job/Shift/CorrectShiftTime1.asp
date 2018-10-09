<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
action="location.href='"&beforepath&"'"

jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
sequence=request.QueryString("sequence")
shifttype=request.QueryString("shifttype")

pagename="/Job/Shift/CorrectShiftTime.asp"
SQL="select SHIFT_IN_TIME,SHIFT_OUT_TIME,SHIFT_IN_PERSON,SHIFT_OUT_PERSON from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	if shifttype="in" then
		new_shift_in_time=""
		new_shift_in_person=""
		a_shift_in_time=split(rs("SHIFT_IN_TIME"),",")
		a_shift_in_person=split(rs("SHIFT_IN_PERSON"),",")
		for i=0 to ubound(a_shift_in_time)-1
			if i<>cint(sequence) then
			new_shift_in_time=new_shift_in_time&a_shift_in_time(i)&","
			new_shift_in_person=new_shift_in_person&a_shift_in_person(i)&","
			end if
		next
		rs("SHIFT_IN_TIME")=new_shift_in_time
		rs("SHIFT_IN_PERSON")=new_shift_in_person
	else
		new_shift_out_time=""
		new_shift_out_person=""
		a_shift_out_time=split(rs("SHIFT_OUT_TIME"),",")
		a_shift_out_person=split(rs("SHIFT_OUT_PERSON"),",")
		for i=0 to ubound(a_shift_out_time)-1
			if i<>cint(sequence) then
			new_shift_out_time=new_shift_out_time&a_shift_out_time(i)&","
			new_shift_out_person=new_shift_out_person&a_shift_out_person(i)&","
			end if
		next
		rs("SHIFT_OUT_TIME")=new_shift_out_time
		rs("SHIFT_OUT_PERSON")=new_shift_out_person
	end if
	rs.update
end if
rs.close
word="已删除选定的停线/开线时间！"
'calculate cycle time again
SQL="select SHIFT_IN_TIME,SHIFT_OUT_TIME,START_TIME,CLOSE_TIME,CYCLE_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	if rs("SHIFT_IN_TIME")<>"" and rs("SHIFT_OUT_TIME")<>"" then
	shift_in_time=left(rs("SHIFT_IN_TIME"),len(rs("SHIFT_IN_TIME"))-1)
	a_shift_in_time=split(shift_in_time,",")
	shift_out_time=left(rs("SHIFT_OUT_TIME"),len(rs("SHIFT_OUT_TIME"))-1)
	a_shift_out_time=split(shift_out_time,",")
		if ubound(a_shift_in_time)=ubound(a_shift_out_time) then
			for i=0 to ubound(a_shift_in_time)
			shift_interval=shift_interval+datediff("n",a_shift_out_time(i),a_shift_in_time(i))
			next
			elapsed_time=cstr(datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))-shift_interval)
			rs("CYCLE_TIME")=elapsed_time
			rs.update
			word=word&"\n周转时间更新成功！"
		else
			word=word&"\n周转时间更新失败！"
		end if
	end if
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>")
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->