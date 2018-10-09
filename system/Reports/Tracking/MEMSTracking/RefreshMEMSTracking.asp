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
id=request.QueryString("id")
SQL="select * from JOB_MEMS_LIST where NID='"&id&"'"
rs.open SQL,conn,1,3
MEMS_name=rs("TRACKING_NAME")
if request("fromdate")="" then
fromdate=rs("WEEK_START_DAY")
else
fromdate=request("fromdate")
end if
if request("todate")="" then
todate=rs("WEEK_END_DAY")
else
todate=request("todate")
end if
rs.close

path=request.QueryString("path")
query=request.QueryString("query")
set rsU=server.CreateObject("adodb.recordset")
memslot=request("memslot")
memspart=request("memspart")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by DJM.MEMS_LOT asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if memslot<>"" then
where=where&" and DJM.MEMS_LOT like '%"&memslot&"%'"
end if
if memspart<>"" then
where=where&" and DJM.MEMS_PART='"&memspart&"'"
end if
weekindex=datepart("ww",cdate(fromdate),2)
pagepara="&schedule_id="&schedule_id&"&part_id="&part_id
pagename="/Reports/Tracking/MEMSTracking/RefreshMEMSTracking.asp"
SQL="truncate table JOB_MEMS_TEMP"
rs.open SQL,conn,1,3
SQL="select 1,JM.* from JOB_MEMS_RAW JM where JM.START_TIME>=TO_DATE('"&fromdate&"','YYYY-MM-DD') and JM.START_TIME<=TO_DATE('"&todate&"','YYYY-MM-DD')"
session("SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof 
		if rs("MEMS_LOT")<>"" then
			a_MEMS_lot=split(rs("MEMS_LOT"),",")
			if rs("MEMS_QUANTITY")<>"" then
			a_MEMS_quantity=split(rs("MEMS_QUANTITY"),",")
			else
				f_MEMS_QUANTITY=""
				for i=0 to ubound(a_MEMS_lot)
				f_MEMS_QUANTITY=f_MEMS_QUANTITY&"0"&","
				next
				f_MEMS_QUANTITY=left(f_MEMS_QUANTITY,len(f_MEMS_QUANTITY)-1)
			a_MEMS_quantity=split(f_MEMS_QUANTITY,",")
			end if
			if ubound(a_MEMS_lot)<>ubound(a_MEMS_quantity) then
			org=ubound(a_MEMS_quantity)
			diff=ubound(a_MEMS_lot)-ubound(a_MEMS_quantity)
			ReDim Preserve a_MEMS_quantity(UBound(a_MEMS_quantity) +diff)
			thiserror=thiserror&"<a href='/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"' target='_blank'>"&rs("JOB_NUMBER")&"</a>; "
			end if
			for i=0 to ubound(a_MEMS_lot)
			SQLU="select MEMS_LOT,MEMS_PART,MEMS_QUANTITY,MACHINE_NUMBER,JOB_NUMBER,STATION_START_TIME,PART_NUMBER_ID,CURRENT_STATION_ID,RND_KEY,CREATOR_CODE from JOB_MEMS_TEMP"
			rsU.open SQLU,conn,3,3
			rsU.addnew
			rsU("MEMS_LOT")=a_MEMS_lot(i)
			rsU("MEMS_PART")=rs("MEMS_PART")
			if a_MEMS_quantity(i)<>"" then
			rsU("MEMS_QUANTITY")=a_MEMS_quantity(i)
			else
			rsU("MEMS_QUANTITY")=0
			end if
			rsU("MACHINE_NUMBER")=rs("MACHINE_NUMBER")
			rsU("JOB_NUMBER")=rs("JOB_NUMBER")
			rsU("STATION_START_TIME")=rs("START_TIME")
			rsU("PART_NUMBER_ID")=rs("PART_NUMBER_ID")
			rsU("CURRENT_STATION_ID")=rs("CURRENT_STATION_ID")
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
SQL="select 1,DJM.MEMS_LOT,DJM.MACHINE_NUMBER,DJM.MEMS_PART from DISTINCT_JOB_MEMS_TEMP DJM where DJM.MEMS_LOT is not null "&where&order
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
<form name="form1" method="post" action="/Reports/Tracking/MEMSTracking/RefreshMEMSTracking.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Refresh MEMS Tracking </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">MEMS Lot Number</span> </td>
    <td height="20"><input name="memslot" type="text" id="memslot" value="<%=memslot%>"></td>
    <td>MEMS Part Number </td>
    <td><input name="memspart" type="text" id="memspart" value="<%=memspart%>"></td>
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
  <td height="20" colspan="24" class="t-c-greenCopy">Refresh MEMS Tracking form <% =formatdate(fromdate,application("shortdateformat"))%> to <% =formatdate(todate,application("shortdateformat"))%></td>
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
  <td height="20" colspan="24" class="strongred"><span title="MEMS Quanity in following jobs has error, please revise it and try again."><span class="strongred">Error jobs:</span> <%= thiserror %></span></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJM.MEMS_LOT&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">MEMS Lot Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJM.MEMS_LOT&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJM.MEMS_PART&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">MEMS PART Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=DJM.MEMS_PART&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Machine NO</div></td>
  <%for j=1 to 5%>
  <td class="t-t-Borrow"><div align="center">Job <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Part <%=j%></div></td>
  <td class="t-t-Borrow"><div align="center">Station Time<%=j%></div></td>
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
    <td height="20"><div align="center"><% =rs("MEMS_LOT") %></div></td>
    <td><div align="center">
        <% =rs("MEMS_PART") %>
    </div></td>
    <td><div align="center"><% =rs("MACHINE_NUMBER") %></div></td>
    <%SQLU="select 1,JM.JOB_NUMBER,JM.STATION_START_TIME,P.PART_NUMBER,JM.CURRENT_STATION_ID,JM.MEMS_QUANTITY from JOB_MEMS_TEMP JM inner join PART P on JM.PART_NUMBER_ID=P.NID where JM.MEMS_LOT='"&rs("MEMS_LOT")&"' and JM.MEMS_PART='"&rs("MEMS_PART")&"' and JM.MACHINE_NUMBER='"&rs("MACHINE_NUMBER")&"' order by JM.JOB_NUMBER"
	rsU.open SQLU,conn,1,3
	if not rsU.eof then
		for j=1 to 5
  			if not rsU.eof then%>
		  <td><div align="center"><% =rsU("JOB_NUMBER")%>&nbsp;</div></td>
			<td><div align="center"><% =rsU("PART_NUMBER")%>&nbsp;</div></td>
			<td><div align="center"><%=rsU("STATION_START_TIME")%></div></td>
			<td><div align="center"><% =rsU("MEMS_QUANTITY")%>&nbsp;</div></td>
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
<form name="form2" method="post" action="/Reports/Tracking/MEMSTracking/RefreshMEMSTracking1.asp">
  <tr>
    <td height="20" colspan="24">Refresh time:
      <% =formatdate(create_time,application("longdateformat"))%>
&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="24">Report name:
      <input name="MEMS_name" type="text" id="MEMS_name" value="<%=MEMS_name%>">
      <input name="yearindex" type="hidden" id="yearindex" value="<%=year(fromdate)%>">
      <input name="weekindex" type="hidden" id="weekindex" value="<%=weekindex%>">
      <input name="fromdate" type="hidden" id="fromdate" value="<%=fromdate%>">
      <input name="todate" type="hidden" id="todate" value="<%=todate%>">
      <input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input name="Save" type="submit" id="Save" value="Update This Report"></td>
  </tr>
</form>
<%else%>
  <tr>
    <td height="20" colspan="24"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->