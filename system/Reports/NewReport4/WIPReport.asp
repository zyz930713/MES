<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
action=request.QueryString("action")
batchno=request("batchno")
modelname=request("modelname")
stationname=request("station")
fromdate=request("fromdate")
fromtime=request("fromtime")
input=0
moutput=0
time0=now   
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

if isnull(fromtime) or fromtime=""  then
	fromtime="14:30:00"
end if
todate=request("todate")
if isnull(todate) or todate=""  then
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
totime=request("totime")
if isnull(totime) or totime="" then
	totime="14:30:00"
end if

if(action="GenereateReport")then
 SQL="SELECT * FROM JOB_RETEST_QA_HISTORY WHERE IS_DELETE='0' AND IS_FINISHED='0' "
 if(batchno<>"")then
 	SQL=SQL+" and BATCHNO='"+batchno+"'"
 end if 
  if(modelname<>"")then
 	SQL=SQL+" and MODEL_NAME='"+modelname+"'"
 end if 
 if(stationname="SLAPPING")then
 	SQL=SQL+" and IS_SPLAPPING='1'"
 end if
 if(stationname="RETEST")then
 	SQL=SQL+" and IS_SPLAPPING='0' AND STATION_NAME='RETEST'"
 end if
  if(stationname="QA")then
 	SQL=SQL+" AND STATION_NAME='QA'"
 end if		
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
		if(form1.station.options(form1.station.selectedIndex).value=="0")
		{
			window.alert("Please select one station!");
			return;
		}

		form1.action="WIPReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="WIPReport.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Slapping Retest QA WIP</td>
  </tr>
  <tr>
    <td height="20">BatchNO </td>
     <td height="20">
	 	 <input type="text" name="batchno" id="batchno" value="<%=batchno%>">
	</td>
    <td>Part Number</td>
     <td height="20"> 
	 	<input type="text" name="modelname" id="modelname" value="<%=modelname%>">
	 </td>
  
    <td>Station Name</td>
    <td height="20">
		<select name="station" id="station">
    		<option value="0">-- Select Station --</option>
	   		<option value="SLAPPING" <%if stationname="SLAPPING" then response.write "selected" end if %>>Slapping</option>
			<option value="RETEST" <%if stationname="RETEST" then response.write "selected" end if %>>Retest</option>
			<option value="QA" <%if stationname="QA" then response.write "selected" end if %>>QA</option>
 	 	</select>  
	 </td>
	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  	<Td><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
  </tr>
  
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-t-Borrow">Batch NO</td>
	 <td height="20" class="t-t-Borrow">Job Number</td>
	 <td height="20"class="t-t-Borrow">Part Number</td>
	  <td height="20" class="t-t-Borrow">Qty</td>
	  <td height="20" class="t-t-Borrow">Operator</td>
	   <td height="20" class="t-t-Borrow">Station</td>
	  <td height="20" class="t-t-Borrow">In Station Time</td>
  </tr>
  <%if action="GenereateReport" then%>
  <% while not rs.eof%>
  	<tr>
    <td height="20" ><%=rs("BATCHNO")%></td>
	 <td height="20" ><%=rs("JOB_NUMBER")%></td>
	 <td height="20" ><%=rs("MODEL_NAME")%></td>
	  <td height="20" ><%=rs("STATION_IN_QTY")%></td>
	   <td height="20" ><%=rs("STATION_IN_OPERATOR")%></td>
	  <td height="20" ><%=rs("STATION_NAME")%></td>
	   <td height="20" ><%=rs("STATION_IN_DATETIME")%></td>
  </tr>
  <% 
  	rs.movenext
  	wend%>
  <%end if%>

</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->