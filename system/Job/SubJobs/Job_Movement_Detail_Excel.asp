<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.Expires=0
response.CacheControl="no-cache"

%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobStation.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
	currentstation=trim(request("currentstation"))
	linename=request("txtLineName")
	
	jobclosetimefromhour=request("txtJobCloseTimeFromHour")
	if(jobclosetimefromhour="") then
		jobclosetimefromhour="00:00:00"
	end if
	JobCloseTimeToHour=request("txtJobCloseTimeToHour")
	if(JobCloseTimeToHour="") then
		JobCloseTimeToHour="23:59:59"
	end if
	
	timeflag=request("dropTimeFlag")
	
	fromdate=request("txtJobCloseTimeFrom")
	todate=request("txtJobCloseTimeTo")
	SQL="select js.job_number,js.sheet_number,PART_NUMBER_TAG,LINE_NAME,station_name,JS.start_time,JS.close_time,JS.Station_Start_quantity,JS.good_quantity "
	SQL=SQL & "from job_stations JS,JOb j,station S"
	SQL=SQL &" where JS.job_number=j.job_number and js.sheet_number=j.sheet_number and JS.station_id=s.NID"
	IF(currentstation <> "") then
		SQL=SQL & " AND s.MOTHER_STATION_ID='"&currentstation&"'"
	end if
	IF(fromdate <> "") then
		if(timeflag="" or timeflag="0") then
			SQL=SQL & " AND JS.close_time>=TO_DATE('"& fromdate &" " & jobclosetimefromhour & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time>=TO_DATE('"& fromdate &" " & jobclosetimefromhour & "','YYYY-MM-DD HH24:MI:SS')"
		end if
	end if
	
	IF(todate <> "") then
		if(timeflag="" or timeflag="0") then
			SQL=SQL & " AND JS.close_time<=TO_DATE('"& todate &" "& JobCloseTimeToHour & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time<=TO_DATE('"& todate &" "& JobCloseTimeToHour & "','YYYY-MM-DD HH24:MI:SS')"
		end if
	end if
	
	if(linename<>"") then
		SQL=SQL & " AND J.line_name='"&linename&"'"
	end if
	
	SQL=SQL & " order by PART_NUMBER_TAG,station_name,LINE_NAME,job_number,sheet_number "
 
	rs.open SQL,conn,1,3

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=application("SystemName")%></title>
	 
</head>
<body>
<form action="" method="post" name="form1">
  
	<%
	  Response.ContentType = "application/vnd.ms_excel" 
	  Response.AddHeader "content-Disposition","attachment;filename=Job_Movement_Detail.xls"
	
	%>
		<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
		<tr>
			<Td class="t-t-Borrow">Job_Number</Td>
			<Td class="t-t-Borrow">Job_Sheet_Number</Td>
			<Td class="t-t-Borrow">Part_Number</Td>
			<Td class="t-t-Borrow">Line_Name</Td>
			<Td class="t-t-Borrow">Station_Name</Td>
			<Td class="t-t-Borrow">Start_Time</Td>
			<Td class="t-t-Borrow">Close_Time</Td>
			<Td class="t-t-Borrow">Station_Start_Qty</Td>
			<Td class="t-t-Borrow">Good_Qty</Td>
		</tr>
		 <%	while not rs.eof %>
			<tr>
				<td><%=rs(0).value%></td>
				<td><%=rs(1).value%></td>
				<td><%=rs(2).value%></td>
				<td><%=rs(3).value%></td>
				<td><%=rs(4).value%></td>
				<td><%=rs(5).value%></td>
				<td><%=rs(6).value%></td>
				<td><%=rs(7).value%></td>
				<td><%=rs(7).value%></td>
			</tr>
		 <% rs.movenext 
		 wend %>
		 </table>
	 </td>
	 </Tr>
	</table>
</body>
</html>

