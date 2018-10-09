<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
server.ScriptTimeout=99999999

%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<table border="1">
<%
set oRs=server.CreateObject("adodb.recordset")
'SQL="select * from series_group"

SQL="select JOB_NUMBER,INPUT_TIME,COMPLETE_DATE,GET_WORK_TIME(INPUT_TIME,COMPLETE_DATE)*24 as AVG_LEADTIME from job_master where to_char(COMPLETE_DATE,'yyyy-mm-dd')<=to_char(to_date(to_char('2008-4-18'),'yyyy-mm-dd'),'yyyy-mm-dd') and  to_char(COMPLETE_DATE,'yyyy-mm-dd')>=to_char(to_date(to_char('2008-4-12'),'yyyy-mm-dd'),'yyyy-mm-dd') and COMPLETE_DATE is not null"

rs.open SQL,conn,1,3

do while not rs.eof

			avg_leadtime=0
			'sSQL="select avg(GET_WORK_TIME(start_time,close_time))*24 as AVG_LEADTIME from job where GET_JOB_FAMILY(PART_NUMBER_TAG)='" & TRIM(rs("SERIES_GROUP_NAME")&"") & "' and  to_char(close_time,'yyyy-mm-dd')<=to_char(to_date(to_char('2008-4-18'),'yyyy-mm-dd'),'yyyy-mm-dd') and  to_char(close_time,'yyyy-mm-dd')>=to_char(to_date(to_char('2008-4-12'),'yyyy-mm-dd'),'yyyy-mm-dd') and status=1"
			'sSQL="select avg(GET_WORK_TIME(INPUT_TIME,COMPLETE_DATE))*24 as AVG_LEADTIME from job_master where GET_JOB_FAMILY(PART_NUMBER_TAG)='" & TRIM(rs("SERIES_GROUP_NAME")&"") & "' and  to_char(COMPLETE_DATE,'yyyy-mm-dd')<=to_char(to_date(to_char('2008-4-18'),'yyyy-mm-dd'),'yyyy-mm-dd') and  to_char(COMPLETE_DATE,'yyyy-mm-dd')>=to_char(to_date(to_char('2008-4-12'),'yyyy-mm-dd'),'yyyy-mm-dd') and COMPLETE_DATE is not null"
			'sSQL="select GET_WORK_TIME(INPUT_TIME,COMPLETE_DATE)*24 as AVG_LEADTIME from job_master where GET_JOB_FAMILY(PART_NUMBER_TAG)='" & TRIM(rs("SERIES_GROUP_NAME")&"") & "' and  to_char(COMPLETE_DATE,'yyyy-mm-dd')<=to_char(to_date(to_char('2008-4-18'),'yyyy-mm-dd'),'yyyy-mm-dd') and  to_char(COMPLETE_DATE,'yyyy-mm-dd')>=to_char(to_date(to_char('2008-4-12'),'yyyy-mm-dd'),'yyyy-mm-dd') and COMPLETE_DATE is not null"
			'oRs.open sSQL,conn,1,3
			'if oRs.eof then
'				avg_leadtime=0
'			else
'				avg_leadtime=trim(oRs("AVG_LEADTIME")&"")
'				
'				if avg_leadtime="" then avg_leadtime=0
'			end if
'			oRs.close
'			response.Write("<tr><td>" & rs("SERIES_GROUP_NAME") & "</td><td>" & avg_leadtime & "</td></tr>") 
			response.Write("<tr><td>" & rs("JOB_NUMBER") & "</td><td>" & rs("INPUT_TIME") & "</td><td>" & rs("COMPLETE_DATE") & "</td><td>" & trim(rs("AVG_LEADTIME")&"") & "</td></tr>") 
	rs.movenext
loop

rs.close

set rs=nothing
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->