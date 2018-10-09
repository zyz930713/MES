<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetScheduledJobsInDay.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
schedule_id=request.QueryString("schedule_id")
schedule_name=trim(request("schedule_name"))
Schedule_report_time=trim(request("schedule_report_time"))
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/WIP/WIPRporte.asp"
SQL="truncate table WIP_DETAIL_TEMP"
rs.open SQL,conn,3,3
SQL="select NID,PART_NUMBER from PART where STATUS=1 "
session("SQL")=SQL
rs.open SQL,conn,1,3
SQLT="select * from JOB_SCHEDULE_LIST where NID='"&schedule_id&"'"
rsT.open SQLT,conn,1,3
week_start_day=rsT("WEEK_START_DAY")
week_end_day=rsT("WEEK_END_DAY")
rsT.close
day_interval=datediff("d",week_start_day,week_end_day)
Tcount=day_interval+3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Job Schedule: "<%=schedule_name%> (<%=schedule_report_time%>)"</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <%
  scheduleday=week_start_day
  for j=0 to day_interval%>
  <td class="t-t-Borrow"><div align="center"><%=formatdate(scheduleday,application("shortdateformat"))%></div></td>
  <%
  scheduleday=dateadd("d",scheduleday,1)
  next%>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('SchedulePartDetail.asp?schedule_id=<%=schedule_id%>&part_id=<%=rs("NID")%>&part_name=<%=rs("PART_NUMBER")%>')"><%= rs("PART_NUMBER") %></span></div></td>
    <%scheduleday=week_start_day
	for j=0 to day_interval%>
	  <div align="center"><td><div align="center"><span style="cursor:hand" onClick="javascript:window.open('ScheduleDetail.asp?schedule_id=<%=schedule_id%>&part_id=<%=rs("NID")%>&start_time=<%=scheduleday%>')"><%=getScheduledJobsInDay(schedule_id,rs("NID"),scheduleday)%></span>&nbsp;</div></td></div>
	<%scheduleday=dateadd("d",scheduleday,1)
	next%>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
      <form name="form1" method="post" action="/Reports/Process/WIP/SaveWIP.asp">
  </form>
<%
else
%>

  <tr>
    <td height="20" colspan="3"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsT=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->