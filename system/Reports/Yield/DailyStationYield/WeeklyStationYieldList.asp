<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetWeeklyStationYieldByDay.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetProfileTask.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
factory_id=request("factory_id")
station=request("station")
set rs1=server.CreateObject("adodb.recordset")
if request("from_year_number")="" or request("from_week_number")="" then
from_year_number=year(date())
from_week_number=DatePart("ww",dateadd("m",-1,date()))
else
from_year_number=cint(request("from_year_number"))
from_week_number=cint(request("from_week_number"))
end if
if request("to_week_number")="" or request("to_year_number")="" then
to_year_number=year(date())
to_week_number=DatePart("ww",date)
else
to_week_number=cint(request("to_week_number"))
to_year_number=cint(request("to_year_number"))
end if
if factory_id<>"" then
where=where&" and S.FACTORY_ID='"&factory_id&"'"
else
where=where&" and S.FACTORY_ID is null"
end if
if station<>"" then
where=where&" and S.NID='"&station&"'"
end if

FactoryRight ""
SQL="select S.NID,S.STATION_NAME,S.FACTORY_ID,F.FACTORY_NAME from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID where S.NID is not null "&where&" order by S.FACTORY_ID,S.STATION_NAME"
rs1.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/Yield/DailyStationYield/formcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Yield/DailyStationYield/WeeklyStationYieldList.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="5" class="t-c-greenCopy">Detail Span Selection </td>
    </tr>
    <tr>
      <td class="t-b-blue">Factory</td>
      <td><select name="factory_id" id="factory_id">
        <option value="">-- Select --</option>
        <%=getFactory("OPTION",factory_id,factorywhereinside,"","")%>
      </select></td>
      <td class="t-b-blue">Station</td>
      <td><select name="station" id="station">
        <option value="">--Select--</option>
        <%FactoryRight "S."%>
        <%= getStation(null,"OPTION",station,factorywhereoutside," order by S.STATION_NAME",null,null) %>
      </select></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="t-b-blue">Week Number</td>
      <td>from
        <input name="from_week_number" type="text" id="from_week_number" value="<%=from_week_number%>" size="2">
&nbsp;of
<input name="from_year_number" type="text" id="from_year_number" value="<%=from_year_number%>" size="4">
to
<input name="to_week_number" type="text" id="to_week_number" value="<%=to_week_number%>" size="2">
&nbsp;of
<input name="to_year_number" type="text" id="to_year_number" value="<%=to_year_number%>" size="4"></td>
      <td colspan="3"><input name="Refresh" type="submit" class="t-b-Yellow" id="Refresh" value="Refresh">
&nbsp;
<input name="ontime" type="button" class="t-t-DarkBlue" id="ontime" value="Ontime Report" onClick="location.href='OnTimeStationYield.asp'">
<input name="daily" type="button" class="t-t-DarkBlue" id="daily" value="Daily Report" onClick="location.href='DailyStationYieldList.asp'"></td>
    </tr>
  </table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" byield="1" byieldcolorlight="#666666" byieldcolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="<%=to_week_number-from_week_number+3%>" class="t-c-greenCopy">
      <table width="100%" cellpadding="0" cellspacing="0" byield="0">
        <tr>
          <td>Weekly Station Yield from <%=from_week_number%> of <%=from_year_number%> to <%=to_week_number%> of <%=to_year_number%>
          <input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Create new weekly report" onClick="location.href='WeeklyStationYield.asp'"></td>
        </tr>
      </table>    </td>
  </tr>
  <tr class="t-b-blue">
    <td width="242" rowspan="2"><div align="center">Factory</div></td>
    <td width="286" rowspan="2"><div align="center">Station Name</div></td>
    <%
	this_year_number=from_year_number
	this_week_number=from_week_number
	while this_year_number<=to_year_number and this_week_number<=to_week_number%>
    <td width="500" height="20"><div align="center"><%=this_week_number%>&nbsp;of <%=this_year_number%>&nbsp;</div></td>
    <%this_week_number=this_week_number+1
	if this_week_number>53 then
	this_week_number=1
	this_year_number=this_year_number+1
	end if
	wend%>
  </tr>
  <tr class="t-b-blue">
  <%
	this_year_number=from_year_number
	this_week_number=from_week_number
	while this_year_number<=to_year_number and this_week_number<=to_week_number%>
    <td height="20"><div align="center"><%GetWeeklyStationYieldSpan this_year_number,this_week_number,from_time,to_time%><%=from_time%> - <%=to_time%>&nbsp;<%if from_time<>"" and to_time<>"" and factory_id<>"" then%><span style="cursor:hand" class="red" onClick="window.open('WeeklyStationYield1.asp?factory_id=<%=factory_id%>&year_index=<%=this_year_number%>&week_number=<%=this_week_number%>&from_time=<%=from_time%>&to_time=<%=to_time%>&report_type=RE')" title="Regenerate">R</span>&nbsp;<span style="cursor:hand" class="red" onClick="location.href('DeleteWeeklyStationYield.asp?factory_id=<%=factory_id%>&year_number=<%=this_year_number%>&week_number=<%=this_week_number%>&path=<%=path%>&query=<%=query%>')" title="Delete">D</span><%end if%></div></td>
	<%this_week_number=this_week_number+1
	if this_week_number>53 then
	this_week_number=1
	this_year_number=this_year_number+1
	end if
	wend%>
  </tr>
  <%
  if not rs1.eof then
  while not rs1.eof%>
  <tr>
    <td><%=rs1("FACTORY_NAME")%></td>
    <td><%=rs1("STATION_NAME")%></td>
	<%
	this_year_number=from_year_number
	this_week_number=from_week_number
	while this_year_number<=to_year_number and this_week_number<=to_week_number
	yield=formatpercent(GetWeeklyStationYieldByDay(this_year_number,this_week_number,rs1("NID"),from_time,to_time),2,-1)%>
    <td height="20" <%if this_day=date() then%>class="t-t-Borrow"<%end if%>><span style="cursor:hand" onClick="window.open('WeeklyStationYieldDetail.asp?station_id=<%=rs1("NID")%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><%=yield%></span></td>
    <%this_week_number=this_week_number+1
	if this_week_number>53 then
	this_week_number=1
	this_year_number=this_year_number+1
	end if
	wend%>
  </tr>
  <%rs1.movenext
  wend
  end if
  rs1.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->