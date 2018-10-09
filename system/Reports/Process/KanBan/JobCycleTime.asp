<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->


<%
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")
jobnumber=request("jobnumber")
input=0
moutput=0
time0=now   
line=request("line")
ModelName=request("ModelName")

if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

if isnull(fromtime) or fromtime=""  then
	fromtime="14:30:00"
end if
todate=request("todate")
if isnull(todate) or todate=""  then
	'todate=cstr(year(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(month(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(day(dateadd("d",7-weekday(time0),time0)))
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
totime=request("totime")
if isnull(totime) or totime="" then
	totime="14:30:00"
end if


'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
		 
	sql="select j.job_number,j.sheet_number,j.part_number_tag,j.job_start_quantity ,j.current_station_id, j.job_good_quantity,"
	sql=sql+" j.start_time, s.STATION_NAME|| '(' || s.STATION_CHINESE_NAME || ')' as current_station, "
	sql=sql+" ROUND(TO_NUMBER(sysdate-j.previous_station_close_time)*24) as current_Station_Cycle_Time,"
	sql=sql+" ROUND(TO_NUMBER(sysdate-j.start_time)*24) as cumulate_Cycle_Time,j.stations_index"
	sql=sql+" from job j, station s "
	sql=sql+" where j.current_station_id=s.nid "
	sql=sql+" and j.status='0' "
	sql=sql+" and close_time is null and start_time>=to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
	sql=sql+" and close_time is null and start_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
	
	if jobnumber<>"" then
		SQL=SQL+" and j.job_number='"+jobnumber+"'"
	end if 
	if line<>"" then
		SQL=SQL+" and j.line_name='"+line+"'"
	end if 
	if ModelName<>"" then
		SQL=SQL+" and j.part_number_tag='"+ModelName+"'"
	end if 
	sql=sql+" ORDER BY j.JOB_NUMBER,j.start_time DESC"
	
	session("SQL")=SQL
	session("FromDateTime") =fromdate & " " &fromtime
	session("ToDateTime") =todate & " " &totime
	rs.open SQL,conn,1,3
	
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	function GenerateReport()
	{
		form1.action="JobCycleTime.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="JobCycleTime.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-b-midautumn">Job Cycle Report</td>
  </tr>

  <Tr>
   <td>DateTime From:</td> 
	 <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	<select name="fromtime" id="fromtime">
   			 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
</td>
<td>To:</td>
<td>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
	</td>
  	<Td width="10%">Job Number</Td>
	<Td width="10%"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" size="16"></Td>
  	
	 
  </Tr>
  <tr>
  	<Td>Line</Td>
	<Td><input name="line" type="text" id="line" value="<%=line%>" size="16"></Td>
	<Td>Model Name</Td>
	<Td><input name="ModelName" type="text" id="ModelName" value="<%=ModelName%>" size="16"></Td>
	 
	<Td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="10"> </td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Sheet Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Model Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Job Start Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Current Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Job Start Time</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Current Station</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Current Station Cycle Time(H)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Overall Cycle Time(H)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Routing</div></td>
   </tr>
	 
  <% if(rs.State > 0 ) then %>
  <% for j=0 to rs.recordcount-1%>
  <tr>
    <td height="20"><div align="center"><%=rs("job_number")%></div></td>
	<td height="20"><div align="center"><%=rs("sheet_number")%></div></td>
	<td height="20"><div align="center"><%=rs("part_number_tag")%></div></td>
	<td height="20"><div align="center"><%=rs("job_start_quantity")%></div></td>
	<td height="20"><div align="center">
	<%
		sql="select * from job_stations where job_number='"&rs("job_number")&"' and sheet_number='"&rs("sheet_number")&"' and station_id='"+rs("current_station_id")+"' "
		set rsStationQty=server.CreateObject("adodb.recordset")
		rsStationQty.open sql,conn,1,3
		StationQty=0
		if rsStationQty.recordcount>0 then
			StationQty=rsStationQty("STATION_START_QUANTITY")
		else
			StationQty=rs("JOB_GOOD_QUANTITY")
		end if 
		response.write StationQty
	%></div></td>
	<td height="20"><div align="center"><%=rs("start_time")%></div></td>
	<td height="20"><div align="center"><%=rs("current_station")%></div></td>
	<td height="20"><div align="center"><%=rs("current_Station_Cycle_Time")%>&nbsp;</div></td>
	<td height="20"><div align="center"><%=rs("cumulate_Cycle_Time")%>&nbsp;</div></td>
	<td height="20" align="left">
	<%
		StationIndex=rs("stations_index")
		ArrStation=split(StationIndex,",")
		StationStr=""
		for m=0 to ubound(ArrStation)
			SQL="SELECT * FROM STATION WHERE NID='"+ArrStation(m)+"'"
			set rsStation=server.CreateObject("adodb.recordset")
			rsStation.open SQL,conn,1,3
			if rsStation.recordcount>0 then
				StationStr=StationStr+rsStation("STATION_NAME")+"("+rsStation("STATION_CHINESE_NAME")+")"+"->"
			end if 
		next 
		if(len(StationStr)>2) then
			StationStr=left(StationStr,len(StationStr)-2)
		end if 
		response.write StationStr
	%></td>
   </tr>
  <%rs.movenext 
  	next%>
  <%end if %>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->