<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
fromdate=request("fromdate")
todate=request("todate")
where=""
if fromdate="" then
fromdate=date()
end if
if todate="" then
todate=date()
end if

if fromdate<>"" then
where=where&" and to_date(J.start_time)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.start_time)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&fromdate="&fromdate&"&todate="&todate
SQL="select a.ACTION_NAME,sum(nvl(ACTION_VALUE,0)) AS QUANTITY from JOB_ACTIONS ja inner join JOB j on ja.job_number=j.job_number and ja.sheet_number=j.sheet_number inner join ACTION A on ja.action_id=a.nid where action_id in ('AC00000989','AC00000990','AC00000959','AC00000965','AC00000967','AC00000976','AC00000977','AC00000978') and j.start_time>to_date('2010-6-9','yyyy-mm-dd hh24:mi:ss') "&where&" group by a.action_name"
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
<!--#include virtual="/Language/Reports/Scrap/ScrapRecords/Lan_ScrapRecords.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post" action="/Job/Scrap/ScrapRecords/ScrapRecords.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchScrapTime"></span></td>
    <td><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; <img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="3" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="3" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="3"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Code"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td><div align="center"><%=rs("ACTION_NAME")%> </div></td>
		<td><div align="center"><%=rs("QUANTITY")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="3"><div align="center"><span id="inner_NoRecords">No Records</span></div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->