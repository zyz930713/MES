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
ic_id=request("ic_id")
ic_name=request("ic_name")
ic_report_time=request("ic_report_time")
iclot=trim(request("iclot"))
icpart=trim(request("icpart"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by DJI.IC_LOT asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if iclot<>"" then
where=where&" and DJI.IC_LOT like '%"&iclot&"%'"
end if
if icpart<>"" then
where=where&" and DJI.IC_PART='"&icpart&"'"
end if
weekindex=datepart("ww",cdate(fromdate),2)
pagepara="&ic_id="&ic_id&"&ic_name="&ic_name&"&ic_report_time="&ic_report_time&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Tracking/ICTracking/ICTracking.asp"
SQL="select 1,DJI.IC_LOT,DJI.MACHINE_NUMBER,DJI.IC_PART from DISTINCT_JOB_IC_DETAIL DJI where DJI.IC_ID='"&ic_id&"' "&where&order
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
<form name="form1" method="post" action="/Reports/Tracking/ICTracking/ICTracking.asp?<%=replace(query,"*","&")%>">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">Search IC Tracking</td>
  </tr>
  <tr>
    <td height="20">IC Lot Number</td>
    <td height="20"><input name="iclot" type="text" id="iclot" value="<%=iclot%>"></td>
    <td>IC Part Number </td>
    <td><input name="icpart" type="text" id="icpart" value="<%=icpart%>"></td>
    <td>Job Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate" value="<% =formatdate(fromdate,application("shortdateformat"))%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate" value="<% =formatdate(todate,application("shortdateformat"))%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
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
  <td height="20" colspan="24" class="t-c-greenCopy">Browse IC Tracking from 
    <% =formatdate(fromdate,application("shortdateformat"))%> to <% =formatdate(todate,application("shortdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="24" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right">
        </div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="24">Report Name: 
    <% =ic_name%> </td>
</tr>
<tr>
  <td height="20" colspan="24">Generating time:
    <% =formatdate(ic_report_time,application("longdateformat"))%>
&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_LOT&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">IC Lot Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_LOT&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">IC PART Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Machine NO</div></td>
  <%for j=1 to 5%>
  <td class="t-t-Borrow"><div align="center">Job <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Part <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Station Time <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Quantity <%=j%></div></td>
  <%next%>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td height="20"><div align="center"><% =rs("IC_LOT") %></div></td>
    <td><div align="center">
        <% =rs("IC_PART") %>
    </div></td>
    <td><div align="center"><% =rs("MACHINE_NUMBER") %></div></td>
    <%SQLU="select 1,JI.JOB_NUMBER,JI.STATION_START_TIME,P.PART_NUMBER,JI.CURRENT_STATION_ID,JI.IC_QUANTITY from JOB_IC_DETAIL JI inner join PART P on JI.PART_NUMBER_ID=P.NID where JI.IC_LOT='"&rs("IC_LOT")&"' and JI.IC_ID='"&ic_id&"' and JI.IC_PART='"&rs("IC_PART")&"' and JI.MACHINE_NUMBER='"&rs("MACHINE_NUMBER")&"' order by JI.JOB_NUMBER"
	rsU.open SQLU,conn,1,3
	if not rsU.eof then
		for j=1 to 5
  			if not rsU.eof then%>
		  <td><div align="center"><a href="/Job/JobDetail.asp?jobnumber=<%=rsU("JOB_NUMBER")%>" target="_blank"><% =rsU("JOB_NUMBER")%></a>&nbsp;</div></td>
			<td><div align="center"><% =rsU("PART_NUMBER")%>&nbsp;</div></td>
			<td><div align="center"><% =rsU("STATION_START_TIME")%>&nbsp;</div></td>
			<td><div align="center"><% =rsU("IC_QUANTITY")%>&nbsp;</div></td>
	<%
			rsU.movenext
			else%>
			<td><div align="center">&nbsp;</div></td>
			<td><div align="center">&nbsp;</div></td>
			<td><div align="center">&nbsp;</div></td>
			<td><div align="center">&nbsp;</div></td>
	<%
			end if
		next
	end if
	rsU.close%>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
<%else%>
  <tr>
    <td height="20" colspan="24"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing
set rsV=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->