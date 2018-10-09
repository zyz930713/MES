<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
Response.Buffer = False
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
 
	
	SQL="select js.job_number,js.sheet_number,PART_NUMBER_TAG,LINE_NAME,station_name,JS.start_time,JS.close_time,JS.Station_Start_quantity,JS.good_quantity "	
	SQL=SQL & "from job_stations JS,JOb j,station S"
	SQL=SQL &" where JS.job_number=j.job_number and js.sheet_number=j.sheet_number and JS.station_id=s.NID"
	IF(request("currentstation") <> "") then
		SQL=SQL & " AND s.MOTHER_STATION_ID='"&request("currentstation")&"'"
	end if
	IF(session("starttime") <> "") then
		if(session("timeflag")="" or session("timeflag")="0") then
			SQL=SQL & " AND JS.close_time>=TO_DATE('"& session("starttime") & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time>=TO_DATE('"& session("starttime") & "','YYYY-MM-DD HH24:MI:SS')"
		end if
	end if
	
	IF(session("endtime") <> "") then
		if(session("timeflag")="" or session("timeflag")="0") then
			SQL=SQL & " AND JS.close_time<=TO_DATE('"& session("endtime") & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time<=TO_DATE('"& session("endtime") & "','YYYY-MM-DD HH24:MI:SS')"
		end if
	end if
	if(request("LineName")<>"") then
		SQL=SQL & " AND J.line_name='"&request("LineName")&"'"
	end if
	if(request("PartNumber")<>"") then
		SQL=SQL & " AND J.PART_NUMBER_TAG='"&request("PartNumber")&"'"
	end if
	SQL=SQL & " order by  job_number,sheet_number"
	
	rs.open SQL,conn,1,3
%>
<html>
<head>
	<link href="/CSS/General.css" rel="stylesheet" type="text/css">
	<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<Tr>
		<Td class="t-t-Borrow">Job_Number</Td>
		<Td class="t-t-Borrow">Sheet_Number</Td>
		<Td class="t-t-Borrow">Part_Number</Td>
		<Td class="t-t-Borrow">Line_Name</Td>
		<Td class="t-t-Borrow">Station_Name</Td>
		<Td class="t-t-Borrow">Station_Start_Time</Td>
		<Td class="t-t-Borrow">Station_End_Time</Td>	
		<Td class="t-t-Borrow">Station_Start_Qty</Td>
		<Td class="t-t-Borrow">Good_Qty</Td>	
	</Tr>
	<%  
		while not rs.eof %>
			<tr>
				<td><%=rs(0).value%></td>
				<td><%=rs(1).value%></td>
				<td><%=rs(2).value%></td>
				<td><%=rs(3).value%></td>
				<td><%=rs(4).value%></td>
				<td><%=rs(5).value%></td>
				<td><%=rs(6).value%></td>
				<td><%=rs(7).value%></td>
				<td><%=rs(8).value%></td>
			</tr>
		 <% 
		 rs.movenext 
		 wend %>
</table>
</body>
</html>
 <!--#include virtual="/WOCF/BOCF_Close.asp" -->
 