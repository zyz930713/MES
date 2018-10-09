<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetDailyFailureRatioByDay.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetProfileTask.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetDefectCode.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
profile_task_id=request("profile_task_id")
seriesgroup=request("seriesgroup")
defectcode=request("defectcode")
station=request("station")
start_day=request("start_day")
close_day=request("close_day")
set rs1=server.CreateObject("adodb.recordset")
if profile_task_id<>"" then
	SQL="select PARAM1 from PROFILE_TASK where NID='"&profile_task_id&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	factory_id=rs("PARAM1")
	end if
	rs.close
	where=where&" and S.FACTORY_ID='"&factory_id&"'"
else
where=where&" and S.FACTORY_ID is null"
factory_id=""
end if
if station<>"" then
where=where&" and D.STATION_ID='"&station&"'"
end if
if defectcode<>"" then
where=where&" and D.NID='"&defectcode&"'"
end if

if start_day="" then
	this_weekday=weekday(date())
	select case this_weekday
	case 1
	start_day=dateadd("d",0,date())
	case 2
	start_day=dateadd("d",-1,date())
	case 3
	start_day=dateadd("d",-2,date())
	case 4
	start_day=dateadd("d",-3,date())
	case 5
	start_day=dateadd("d",-4,date())
	case 6 'friday
	start_day=dateadd("d",-5,date())
	case 7 'saturday
	start_day=dateadd("d",-6,date())
	end select
else
start_day=cdate(start_day)
end if

if close_day="" then
	this_weekday=weekday(date())
	select case this_weekday
	case 1
	close_day=dateadd("d",6,date())
	case 2
	close_day=dateadd("d",5,date())
	case 3
	close_day=dateadd("d",4,date())
	case 4
	close_day=dateadd("d",3,date())
	case 5
	close_day=dateadd("d",2,date())
	case 6 'friday
	close_day=dateadd("d",1,date())
	case 7 'saturday
	close_day=dateadd("d",0,date())
	end select
else
close_day=cdate(close_day)
end if

FactoryRight ""
SQL="select D.NID,D.DEFECT_NAME,D.STATION_ID,S.STATION_NAME,S.FACTORY_ID,F.FACTORY_NAME from DEFECTCODE D inner join STATION S on D.STATION_ID=S.NID inner join FACTORY F on S.FACTORY_ID=F.NID where S.NID is not null "&where&" order by S.FACTORY_ID,S.STATION_NAME"
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
<script language="JavaScript" src="/Reports/Process/Failure/formcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Process/Failure/DailyFailureRatioList.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="9" class="t-c-greenCopy">Detail Span Selection </td>
    </tr>
    <tr>
      <td class="t-b-blue">Task Type <span class="red">*</span> </td>
      <td><select name="profile_task_id" id="profile_task_id">
        <option value="">--Select--</option>
        <%= getProfileTask("OPTION",profile_task_id," and (lower(TASK_NAME) like '%daily failure ratio%')"," order by TASK_NAME","") %>
      </select></td>
      <td class="t-b-blue">Family<span class="red"> *</span></td>
      <td><select name="seriesgroup" id="seriesgroup">
        <option value="">--Select--</option>
        <%FactoryRight "S."%>
        <%= getSeriesGroup("OPTION",seriesgroup,factorywhereoutside," order by S.SERIES_GROUP_NAME",null) %>
      </select></td>
      <td class="t-b-blue">Station</td>
      <td><select name="station" id="station">
        <option value="">--Select--</option>
		<%FactoryRight "S."%>
        <%= getStation(null,"OPTION",station,factorywhereoutside," order by S.STATION_NAME",null,null) %>
      </select></td>
      <td class="t-b-blue">Days</td>
      <td>From
        <input name="start_day" type="text" id="start_day" value="<%=start_day%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.start_day.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
        &nbsp;to
<input name="close_day" type="text" id="close_day" value="<%=close_day%>" size="10">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.close_day.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp;</td>
      <td><input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Refresh">
        &nbsp;
      <input name="ontime" type="button" class="t-t-DarkBlue" id="ontime" value="Ontime Report" onClick="location.href='OnTimeFailureRatio.asp'">
      &nbsp;
      <input name="weekly" type="button" class="t-t-DarkBlue" id="weekly" value="Weekly Report" onClick="location.href='WeeklyFailureRatioList.asp'"></td>
    </tr>
    <tr>
      <td class="t-b-blue">Defectcode</td>
      <td colspan="8"><select name="defectcode" id="defectcode">
        <option value="">--Select--</option>
		<%FactoryRight "D."%>
        <%= getDefectCode("OPTION",defectcode,factorywhereoutside," order by DEFECT_NAME",null) %>
      </select></td>
    </tr>
  </table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" byield="1" byieldcolorlight="#666666" byieldcolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="<%=datediff("d",cdate(start_day),cdate(close_day))+4%>" class="t-c-greenCopy">
      <table width="100%" cellpadding="0" cellspacing="0" byield="0">
        <tr>
          <td>Daily Failure Ratio in <%=datediff("d",cdate(start_day),cdate(close_day))+1%> days</td>
        </tr>
      </table>    </td>
  </tr>
  <tr class="t-b-blue">
    <td width="173"><div align="center">Factory</div></td>
    <td width="515"><div align="center">Station Name</div></td>
    <td width="344"><div align="center">Defectcode Name</div></td>
    <%
	this_day=start_day
	while this_day<=close_day%>
    <td width="344" height="20" class="<%if weekday(this_day)=1 then%>red<%elseif weekday(this_day)=7 then%>green<%end if%>"><div align="center"><%=this_day%>&nbsp;(<%=shortweekdayconvert(weekday(this_day))%>)&nbsp;<span style="cursor:hand" class="red" onClick="window.open('RegenerateDailyFailureRatioReport.asp?profile_task_id=<%=profile_task_id%>&factory_id=<%=factory_id%>&from_time=<%=this_day%>&to_time=<%=dateadd("d",1,this_day)%>')" title="Regenerate">R</span></div></td>
    <%this_day=dateadd("d","1",this_day)
	wend%>
  </tr>
  <%
  if not rs1.eof then
  while not rs1.eof%>
  <tr>
    <td><%=rs1("FACTORY_NAME")%></td>
    <td><%=rs1("STATION_NAME")%></td>
	<td <%if this_day=date() then%>class="t-t-Borrow"<%end if%>><%=rs1("DEFECT_NAME")%></td>
	<%
	this_day=start_day
	while this_day<=close_day%>
    <td height="20" <%if this_day=date() then%>class="t-t-Borrow"<%end if%>><span style="cursor:hand" onClick="window.open('DailyFailureRatioDetail.asp?station_id=<%=rs1("STATION_ID")%>&defectcode_id=<%=rs1("NID")%>&report_day=<%=this_day%>')"><%=formatpercent(GetDailyFailureRatioByDay(seriesgroup,rs1("NID"),rs1("STATION_ID"),this_day),2,-1)%></span></td>
	<%this_day=dateadd("d","1",this_day)
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