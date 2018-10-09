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
set rsU=server.CreateObject("adodb.recordset")
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
iclot=request("iclot")
icpart=request("icpart")
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
if fromdate="" then
fromdate=formatdate(startday,application("shortdateformat"))
end if
if todate="" then
todate=formatdate(endday,application("shortdateformat"))
else
end if
weekindex=datepart("ww",cdate(fromdate),2)
pagepara="&schedule_id="&schedule_id&"&part_id="&part_id
pagename="/Reports/Tracking/ICTracking/ICTracking.asp"
SQL="truncate table JOB_IC_TEMP"
rs.open SQL,conn,1,3
SQL="select 1,JI.* from JOB_IC_RAW JI where JI.START_TIME>=TO_DATE('"&fromdate&"','YYYY-MM-DD') and JI.START_TIME<=TO_DATE('"&todate&"','YYYY-MM-DD')"
session("SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
	thiserror=""
	while not rs.eof 
		if rs("IC_LOT")<>"" then
			a_IC_lot=split(rs("IC_LOT"),",")
			if rs("IC_QUANTITY")<>"" then
			a_IC_quantity=split(rs("IC_QUANTITY"),",")
			else
				f_IC_QUANTITY=""
				for i=0 to ubound(a_IC_lot)
				f_IC_QUANTITY=f_IC_QUANTITY&"0"&","
				next
				f_IC_QUANTITY=left(f_IC_QUANTITY,len(f_IC_QUANTITY)-1)
			a_IC_quantity=split(f_IC_QUANTITY,",")
			end if
			if ubound(a_IC_lot)<>ubound(a_IC_quantity) then
			org=ubound(a_IC_quantity)
			diff=ubound(a_IC_lot)-ubound(a_IC_quantity)
			ReDim Preserve a_IC_quantity(UBound(a_IC_quantity) +diff)
			thiserror=thiserror&"<a href='/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"' target='_blank'>"&rs("JOB_NUMBER")&"</a>; "
			end if
			for i=0 to ubound(a_IC_lot)
			SQLU="select IC_LOT,IC_PART,IC_QUANTITY,MACHINE_NUMBER,JOB_NUMBER,STATION_START_TIME,PART_NUMBER_ID,CURRENT_STATION_ID,RND_KEY,CREATOR_CODE from JOB_IC_TEMP"
			rsU.open SQLU,conn,3,3
			rsU.addnew
			rsU("IC_LOT")=a_IC_lot(i)
			rsU("IC_PART")=rs("IC_PART")
			if isnumeric(a_IC_quantity(i)) and a_IC_quantity(i)<>"" then
			rsU("IC_QUANTITY")=a_IC_quantity(i)
			else
			rsU("IC_QUANTITY")=0
			thiserror=thiserror&"<a href='/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"' target='_blank'>"&rs("JOB_NUMBER")&"</a>; "
			end if
			rsU("MACHINE_NUMBER")=rs("MACHINE_NUMBER")
			rsU("JOB_NUMBER")=rs("JOB_NUMBER")
			rsU("STATION_START_TIME")=rs("START_TIME")
			rsU("PART_NUMBER_ID")=rs("PART_NUMBER_ID")
			rsU("CURRENT_STATION_ID")=rsU("CURRENT_STATION_ID")
			rsU("RND_KEY")=rnd_key
			rsU("CREATOR_CODE")=session("code")
			rsU.update
			rsU.close
			next
		end if
	rs.movenext
	wend
end if
rs.close
SQL="select 1,DJI.IC_LOT,DJI.MACHINE_NUMBER,DJI.IC_PART from DISTINCT_JOB_IC_TEMP DJI where DJI.IC_LOT is not null "&where&order
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
 <form name="form1" method="post" action="/Reports/Tracking/ICTracking/GenerateICTracking.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search IC Tracking </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">IC Lot Number</span> </td>
    <td height="20"><input name="iclot" type="text" id="iclot" value="<%=iclot%>"></td>
    <td>IC Part Number </td>
    <td><input name="icpart" type="text" id="icpart" value="<%=icpart%>"></td>
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
  <td height="20" colspan="44" class="t-c-greenCopy">Browse IC Tracking form <% =formatdate(fromdate,application("shortdateformat"))%> to <% =formatdate(todate,application("shortdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="44" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right">
        </div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="44" class="strongred">Error jobs: </td>
</tr>
<tr>
  <td height="20" colspan="44"><%= thiserror %>&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_LOT&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">IC Lot Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_LOT&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">IC Part Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJI.IC_PART&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Machine NO</div></td>
  <%for j=1 to 10%>
  <td class="t-t-Borrow"><div align="center">Job <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Part <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Station Time  <%=j%></div></td>
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
    <%SQLU="select 1,JI.JOB_NUMBER,JI.STATION_START_TIME,P.PART_NUMBER,JI.CURRENT_STATION_ID,JI.IC_QUANTITY from JOB_IC_TEMP JI inner join PART P on JI.PART_NUMBER_ID=P.NID where JI.IC_LOT='"&rs("IC_LOT")&"' and JI.IC_PART='"&rs("IC_PART")&"' and JI.MACHINE_NUMBER='"&rs("MACHINE_NUMBER")&"' order by JI.JOB_NUMBER"
	rsU.open SQLU,conn,1,3
	if not rsU.eof then
		for j=1 to 10
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
<form name="form2" method="post" action="/Reports/Tracking/ICTracking/SaveICTracking.asp">
  <tr>
    <td height="20" colspan="44">Generating time:
      <% =formatdate(create_time,application("longdateformat"))%>
&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="44">Report name:
      <input name="IC_name" type="text" id="IC_name">
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
    <td height="20" colspan="44"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->