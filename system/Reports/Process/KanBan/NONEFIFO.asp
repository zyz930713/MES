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
		 
	sql="select DISTINCT J.JOB_NUMBER,J.SHEET_NUMBER,J.PART_NUMBER_TAG AS MODEL_NAME ,J.LINE_NAME,J.JOB_START_QUANTITY,J.START_TIME,"
	sql=sql+"  S.STATION_NAME AS JUMP_STATION ,AJS.CHANGE_PERSON ,AJS.CHANGE_REASON AS REASON,AJS.CHANGE_DATETIME"
	sql=sql+" from JOB J ,ABNORMAL_JOB_STATION AJS, STATION S"
	sql=sql+" WHERE J.JOB_NUMBER=AJS.JOB_NUMBER AND J.SHEET_NUMBER=AJS.SHEET_NUMBER"
	sql=sql+" AND AJS.STATION_ID=S.MOTHER_STATION_ID"
	
	SQL=SQL+" and AJS.CHANGE_DATETIME>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
	SQL=SQL+" and AJS.CHANGE_DATETIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		
	if jobnumber<>"" then
		SQL=SQL+" and j.job_number='"+jobnumber+"'"
	end if 
	sql=sql+" ORDER BY AJS.CHANGE_DATETIME DESC,J.JOB_NUMBER"
	
	session("SQL")=SQL
	session("FromDateTime") =fromdate & " " &fromtime
	session("ToDateTime") =todate & " " &totime
	'response.write sql
	 
	
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
		form1.action="NONEFIFO.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="NONEFIFO.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">NONE FIFO Report</td>
  </tr>

  <Tr>
  	<Td>Job Number</Td>
	<Td><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" size="16"></Td>
  	 <td>Job Order DateTime From:</td> 
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
	
&nbsp;
  	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
	
  	<Td><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="18"> </td>
  </tr>
  <tr>
	<%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
    <td height="20" class="t-t-Borrow"><div align="center"><%=rs.Fields(i).name%> </div></td>
  	<%
			next 
		end if
	%>
</tr>
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	<%			
				for i=0 to rs.Fields.count-1
	%>
				 <td height="20" ><div align="center">
				 <%
				 	if(rs.Fields(i).name="SHEET_NUMBER_1" or rs.Fields(i).name="JOB_NUMBER_1") then
						response.write "<a href='../../../job/SubJobs/JobDetail.asp?jobnumber="&rs("Job_Number").value&"&sheetnumber="&rs("sheet_number").value&"' target='_blank'>"
					end if
				 %>
				 <%=rs(i).value%>
				 
				 <%
				 	if(rs.Fields(i).name="SHEET_NUMBER_1" or rs.Fields(i).name="JOB_NUMBER_1") then
						response.write "</a>"
					end if
				 %>
				  
				  </div></td>
				 
				 <%
					if(rs.Fields(i).name="INPUT") then
						input=input+cdbl(rs(i).value)
					end if
					
					if(rs.Fields(i).name="OUTPUT") then
						moutput=moutput+cdbl(rs(i).value)
					end if
					
					if(rs.Fields(i).name="FAMILY_NAME") then
							mFamily=rs(i).value
					end if
					if(rs.Fields(i).name="SERIES_NAME") then
							mSeries=rs(i).value
					end if
					if(rs.Fields(i).name="SUBSERIES_NAME") then
							mSubSeries=rs(i).value
					end if
					if(rs.Fields(i).name="MODEL_NUMBER") then
							mModel=rs(i).value
					end if
				
				 %>
  	<%			
				next
				rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
   <tr>
   <%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
				<% if(rs.Fields(i).name="INPUT") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=input%></div></td>
				<% elseif(rs.Fields(i).name="OUTPUT") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=moutput%></div></td>
				<% elseif(rs.Fields(i).name="FPY") then %>
					<td height="20" class="t-t-Borrow"><div align="center">
							<% if input=0 then response.write "0" else response.write ROUND(moutput/input *100,2) %>%</div></td>
				<% else %>
				<td height="20" class="t-t-Borrow"><div align="center"> &nbsp;</div></td>
				<%end if%>
  	<%     
		    next 
		end if
	%>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->