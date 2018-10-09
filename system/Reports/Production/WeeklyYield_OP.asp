 
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
family=request("family")
Series=request("Series")
SubSeries=request("SubSeries")
model=request("model")
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")

StationName=request("StationName")
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

'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
		SQL="select operator_code,sum(good_quantity) AS OUTPUT,sum(station_start_quantity) AS INPUT,"
		SQL=SQL+" to_char(round((SUM(good_quantity) / sum(STATION_START_QUANTITY)),4)*100) || '%' as Yield , "
		SQL=SQL+" 'WK'||to_char(CLOSE_TIME,'ww') as week from JOB_STATIONS "
		SQL=SQL+" where close_time>=to_date('"+ fromdate+" "+fromtime +"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and   close_time<=to_date('"+ todate+" "+totime +"','yyyy-mm-dd hh24:mi:ss') "
		SQL=SQL+" group by operator_code,TO_CHAR(CLOSE_TIME,'ww')"
		SQL=SQL+" order by operator_code,TO_CHAR(CLOSE_TIME,'ww')"
 
	session("SQL")=SQL
	session("FromDateTime") =fromdate & " " &fromtime
	session("ToDateTime") =todate & " " &totime
	'response.write sql
	'RESPONSE.END
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
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>	
	function GenerateReport()
	{
		 
		form1.action="WeeklyYield_OP.asp?Action=GenereateReport"
		form1.submit();
	}
	
	 
</script>

</head>

<body onLoad="language(<%=session("language")%>);">
<form action="WeeklyYield_OP.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
   
  <Tr align="center">
  	 <td height="20" width="120"><span id="inner_StationCloseTime"> </td> 
	 <td width="500"> <span id="inner_SearchFrom"></span>&nbsp; 
	 <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
	
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
	&nbsp;<span id="inner_SearchTo"></span>:
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
	
&nbsp;
  	<Td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>  	
  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span id="inner_BrowseData"></span></td>        
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right">
      <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('../NewReport/Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div>
	  </td>
    </tr>
  </table></td>
</tr>
<%if(rs.State > 0 ) then%>
  <tr>
  	<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_OpCode"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Input"></span></div></td>	
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Output"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Yield"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Week"></span></div></td>
</tr>
	<%for j=0 to rs.recordcount-1
	
	%>
	<tr align="center">
		<td><%=(j+1)%></td>
		<td height="20" ><%=rs("operator_code")%></td>
		<td height="20" ><%=rs("INPUT")%></td>
		<td height="20" ><%=rs("OUTPUT")%></td>
		<td height="20" ><%=rs("YIELD")%></td>
		<td height="20" ><%=rs("WEEK")%></td>	
	</tr>
	<%rs.movenext
	next 
end if
	%>
</table>

<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->