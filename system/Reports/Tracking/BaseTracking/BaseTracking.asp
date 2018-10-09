<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobTracking/JobTrackingCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
set rsV=server.CreateObject("adodb.recordset")
base_id=request("base_id")
base_name=request("base_name")
base_report_time=request("base_report_time")
baselot=trim(request("baselot"))
basepart=trim(request("basepart"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JB.JOB_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if iclot<>"" then
where=where&" and JB.BASE_LOT like '%"&iclot&"%'"
end if
if icpart<>"" then
where=where&" and JB.BASE_PART='"&icpart&"'"
end if
weekindex=datepart("ww",cdate(fromdate),2)
pagepara="&base_id="&base_id&"&base_name="&base_name&"&base_report_time="&base_report_time&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Tracking/BaseTracking/BaseTracking.asp"
SQL="select 1,JB.*,P.PART_NUMBER from JOB_BASE_DETAIL JB inner join PART P on JB.PART_NUMBER_ID=P.NID where JB.BASE_ID='"&base_id&"' "&where&order
rs.open SQL,conn,1,3
%>
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
<form name="form1" method="post" action="/Reports/Tracking/BaseTracking/BaseTracking.asp?<%=replace(query,"*","&")%>">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">Search Base Tracking</td>
  </tr>
  <tr>
    <td height="20">Base Lot Number</td>
    <td height="20"><input name="baselot" type="text" id="baselot" value="<%=baselot%>"></td>
    <td>Base Part Number </td>
    <td><input name="basepart" type="text" id="basepart" value="<%=basepart%>"></td>
    <td>Job Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate" value="<% =formatdate(fromdate,application("shortdateformat"))%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.form1.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate" value="<% =formatdate(todate,application("shortdateformat"))%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.form1.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy">Browse Base Tracking from 
    <% =formatdate(fromdate,application("shortdateformat"))%> to <% =formatdate(todate,application("shortdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right">
        </div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="18">Report Name: 
    <% =base_name%> </td>
</tr>
<tr>
  <td height="20" colspan="18">Generating time:
    <% =formatdate(base_report_time,application("longdateformat"))%>
&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number </div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Start Time </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Base Part Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td height="20" class="t-t-Borrow" colspan="10"><div align="center">Base Lot Number</div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><a href="/Job/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank">
      <% =rs("JOB_NUMBER")%>
    </a></div></td>
    <td><div align="center">
      <% =rs("PART_NUMBER")%>
    </div></td>
    <td><div align="center">
      <% =rs("START_TIME")%>
    </div></td>
    <td><div align="center">
      <% =rs("BASE_PART") %>
    </div></td>
	<%
	for j=1 to 10
		 if not isnull(rs("BASE_LOT")) and rs("BASE_LOT")<>"" then
		 a_base_lot=split(rs("BASE_LOT"),",")
		 end if%>
	  <td><div align="center">
	  	<%if j-1<ubound(a_base_lot) then%>
		<% =a_base_lot(j-1) %>
		<%end if%>
	  &nbsp;
	<%next%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
<%else%>
  <tr>
    <td height="20" colspan="18"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->