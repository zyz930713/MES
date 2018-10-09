<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobTracking/JobTrackingCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
create_time=now()
rnd_key=gen_key(10)

path=request.QueryString("path")
query=request.QueryString("query")
thisday=date()
select case weekday(thisday,2)
case 1
startday=thisday
endday=dateadd("d",thisday,6)
case 2
startday=dateadd("d",thisday,-1)
endday=dateadd("d",thisday,5)
case 3
startday=dateadd("d",thisday,-2)
endday=dateadd("d",thisday,4)
case 4
startday=dateadd("d",thisday,-3)
endday=dateadd("d",thisday,3)
case 5
startday=dateadd("d",thisday,-4)
endday=dateadd("d",thisday,2)
case 6
startday=dateadd("d",thisday,-5)
endday=dateadd("d",thisday,1)
case 7
startday=dateadd("d",thisday,-6)
endday=thisday
end select
iclot=request("memslot")
icpart=request("memspart")
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by DJB.BOARD_LOT asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if memslot<>"" then
where=where&" and DJB.BOARD_LOT like '%"&memslot&"%'"
end if
if memspart<>"" then
where=where&" and DJB.BOARD_PART='"&memspart&"'"
end if
if fromdate="" then
fromdate=formatdate(startday,application("shortdateformat"))
end if
if todate="" then
todate=formatdate(endday,application("shortdateformat"))
else
end if
weekindex=datepart("ww",cdate(fromdate),2)
pagepara="&schedule_id="&schedule_id&"&part_id="&part_id
pagename="/Reports/Tracking/BOARDTracking/BOARDTracking.asp"
SQL="select 1,JB.*,P.PART_NUMBER from JOB_BOARD_RAW JB inner join PART P on JB.PART_NUMBER_ID=P.NID where JB.START_TIME>=TO_DATE('"&fromdate&"','YYYY-MM-DD') and JB.START_TIME<=TO_DATE('"&todate&"','YYYY-MM-DD')"
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLBOARD "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
 <form name="form1" method="post" action="/Reports/Tracking/BoardTracking/GenerateBoardTracking.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search Board Tracking </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Board Lot Number</span> </td>
    <td height="20"><input name="boardlot" type="text" id="boardlot" value="<%=boardlot%>"></td>
    <td>Board Part Number </td>
    <td><input name="boardpart" type="text" id="boardpart" value="<%=boardpart%>"></td>
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
  <td height="20" colspan="38" class="t-c-greenCopy">Browse Board Tracking form <% =formatdate(fromdate,application("shortdateformat"))%> to <% =formatdate(todate,application("shortdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="38" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right">
        </div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="38" class="strongred">Error jobs: </td>
</tr>
<tr>
  <td height="20" colspan="38"><%= thiserror %>&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number </div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJB.BOARD_LOT&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Board Lot Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJB.BOARD_LOT&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJB.BOARD_PART&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Board Part Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJB.BOARD_PART&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Start Time</div></td>
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
    <td height="20"><div align="center"><% =rs("BOARD_LOT") %></div></td>
    <td><div align="center">
        <% =rs("BOARD_PART") %>
    </div></td>
		  <td><div align="center"><% =rs("START_TIME")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
<form name="form2" method="post" action="/Reports/Tracking/BoardTracking/SaveBoardTracking.asp">
  <tr>
    <td height="20" colspan="38">Generating time:
      <% =formatdate(create_time,application("longdateformat"))%>
&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="38">Report name:
      <input name="BOARD_name" type="text" id="BOARD_name">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
<input name="yearindex" type="hidden" id="yearindex" value="<%=year(fromdate)%>">
      <input name="weekindex" type="hidden" id="weekindex" value="<%=weekindex%>">
      <input name="fromdate" type="hidden" id="fromdate" value="<%=fromdate%>">
      <input name="todate" type="hidden" id="todate" value="<%=todate%>">
      <input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
      <input name="Save" type="submit" id="Save" value="Save This Report"></td>
  </tr>
</form>
<%else%>
  <tr>
    <td height="20" colspan="38"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->