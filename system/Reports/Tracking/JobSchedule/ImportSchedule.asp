<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetJobStationGoodQuantity.asp" -->
<!--#include virtual="/Functions/GetJobStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=todate('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Tracking/JobSchedule/ScheduleList.asp"
SQL="select * from JOB_SCHEDULE_LIST"
session("SQL")=SQL
rs.open SQL,conn,1,3
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
<form name="form1" enctype="multipart/form-data" method="post" action="/Reports/Tracking/JobSchedule/ImportSchedule1.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy">Import Schedule </td>
  </tr>
  <tr>
    <td width="13%" height="20"><span class="style1">Year</span></td>
    <td width="87%" height="20"><input name="yearindex" type="text" id="yearindex" value="<%=year(date())%>" size="6"></td>
  </tr>
  <tr>
    <td height="20">Week Index </td>
    <td height="20">
	  <input name="weekindex" type="text" id="weekindex" value="" size="4">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.weekindex.value=week_index(year + '-' + month + '-' + date)
	document.all.weekstartday.value=year + '-' + month + '-' + date
	document.all.weekendday.value=a_week(year + '-' + month + '-' + date)
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script>
	<script language="vbscript">
	function week_index(thisdate)
	week_index=datepart("ww",cdate(thisdate),2)
	end function
	function a_week(thisdate)
	a_week=formatdatetime(dateadd("d",cdate(thisdate),6),2)
	end function
	</script>
&nbsp; </td>
  </tr>
  <tr>
    <td height="20">Week Start Day </td>
    <td height="20">
      <input name="weekstartday" type="text" id="weekstartday" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.weekstartday.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
</td>
  </tr>
  <tr>
    <td height="20">Week End Day </td>
    <td height="20"><input name="weekendday" type="text" id="weekendday" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.weekendday.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
&nbsp; </td>
  </tr>
  <tr>
    <td height="20">Schedule Name </td>
    <td height="20"><input name="schedulename" type="text" id="schedulename"></td>
  </tr>
  <tr>
    <td height="20">Schedule File </td>
    <td height="20">
      <input type="file" name="file">
    </td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
    </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input type="submit" name="Submit" value="Import">
      &nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->