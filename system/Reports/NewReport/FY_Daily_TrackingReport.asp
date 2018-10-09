<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
 
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")
input=0
moutput=0
time0=now   
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

if isnull(fromtime) or fromtime=""  then
	fromtime="23:59:59"
end if
todate=request("todate")
if isnull(todate) or todate=""  then
	todate=cstr(year(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(month(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(day(dateadd("d",7-weekday(time0),time0)))
end if
totime=request("totime")
if isnull(totime) or totime="" then
	totime="23:59:59"
end if


if action="GenereateReport" then
		sql=" select  distinct jms.job_number,"
		sql=sql+" (select sum(input_quantity) from job_master_store_pre where job_number=jms.job_number and store_type='N' and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as input_quantity,"
		sql=sql+" (select sum(INSPECT_QUANTITY) from job_master_store_pre where job_number=jms.job_number and store_type='N' and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as N_In_Store_Qty,"
		sql=sql+" (select sum(INSPECT_QUANTITY) from job_master_store_pre where job_number=jms.job_number and store_type='S'and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as S_In_Store_Qty,"
		sql=sql+" (select sum(INSPECT_QUANTITY) from job_master_store_pre where job_number=jms.job_number and store_type='R' and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as R_In_Store_Qty,"
		sql=sql+" (select sum(INSPECT_QUANTITY) from job_master_store_pre where job_number=jms.job_number and store_type='C' and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as C_In_Store_Qty,"
		'sql=sql+" (jm.start_quantity)*(1-ss.target_yield/100) as Plan_Rework_Qty,"
		sql=sql+" (select sum(INSPECT_QUANTITY) from job_master_store_pre where job_number=jms.job_number and store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as  Output_WTD,"
		sql=sql+" (select sum(SCRAP_QUANTITY) from JOB_MASTER_SCRAP where job_number=jms.job_number and SCRAP_TIME>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss') and SCRAP_TIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')) as final_scrap_quantity,"
		sql=sql+" ss.target_yield,ss.subseries_name, sn.SERIES_NAME, f.FAMILY_NAME "
		sql=sql+" from job_master_store_pre¡¡jms ,product_model pm,subseries ss,SERIES_NEW sn ,family f"
		sql=sql+" where jms.part_number_tag=pm.item_name"
		sql=sql+" and pm.series_id=ss.nid and pm.series_group_id=sn.nid and pm.family_id=f.nid "
		SQL=SQL+" and jms.store_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and jms.store_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		'SQL=SQL+" and STORE_TYPE='N'"
		sql=sql+" order by ss.subseries_name,jms.job_number"
		session("SQL")=sql
	 
		 
	rs.open sql,conn,1,3
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
	
		form1.action="FY_Daily_TrackingReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="FY_Daily_TrackingReport.asp" method="post" name="form1" id="form1">
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">FY_Daily_Tracking Report</td>
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
	<input name="fromtime" type="text" id="fromtime" value="<%=fromtime%>" size="6">
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
	<input name="totime" type="text" id="totime" value="<%=totime%>" size="6">
&nbsp;
  	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  	<Td><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FY_Daily_Tracking_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
	<Td>&nbsp;</Td>
	<Td>&nbsp;</Td>
	 
  </Tr>
</table>
</form>
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Job_Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Job_Start_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">N_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">S_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">R_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">C_In_Store_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Plan_Rework_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Output_WTD</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Projected_FG_Scrap_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Scrap_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">FG_Scrap_overrun</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">FPY(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Yield(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final_Yield_Target(%)</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">SubSeries_Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Series_Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Family_Name</div></td>
  </tr>
  <%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	<td><%=rs("job_number")%></td>
	<td><%=rs("input_quantity")%></td>
	<td><%=rs("N_In_Store_Qty")%></td>
	<td><%=rs("S_In_Store_Qty")%></td>
	<td><%=rs("R_In_Store_Qty")%></td>
	<td><%=rs("C_In_Store_Qty")%></td>
	<%
		IF ISNULL(rs("N_In_Store_Qty")) THEN
			N_In_Qty=0
		else
			N_In_Qty=rs("N_In_Store_Qty")
		END IF 
		
		IF ISNULL(rs("input_quantity")) THEN
			input_quantity=0
		else
			input_quantity=rs("input_quantity")
		END IF 
		
		IF ISNULL(rs("Output_WTD")) THEN
			Output_WTD=0
		else
			Output_WTD=rs("Output_WTD")
		END IF 
		
		IF ISNULL(rs("final_scrap_quantity")) THEN
			final_scrap_quantity=0
		else
			final_scrap_quantity=rs("final_scrap_quantity")
		END IF 
		
		Plan_Rework_Qty=round(cdbl(input_quantity)*cdbl(rs("target_yield")/100)-cdbl(N_In_Qty),2)
	%>
	<td><%=Plan_Rework_Qty%></td>
	<td><%=rs("Output_WTD")%></td>
	<td><%=cdbl(input_quantity)-cdbl(Output_WTD)%></td>
	<td><%=final_scrap_quantity%></td>
	<td><%=cdbl(final_scrap_quantity)-(cdbl(input_quantity)-cdbl(Output_WTD))%></td>
	<%
		if(cdbl(input_quantity)<>0)then
			fpy=cstr(round((cdbl(N_In_Qty)/cdbl(input_quantity))*100,2))
		else
			fpy=0
		end if 
	%>
	<td><%=fpy%></td>
	<%
		if(cdbl(input_quantity)<>0)then
			finalyield=cstr(round((cdbl(Output_WTD)/cdbl(input_quantity))*100,2))
		else
			finalyield=0
		end if 
	%>
	<td><%=finalyield%></td>
	<td><%=rs("target_yield")%></td>
	<td><%=rs("subseries_name")%></td>
	<td><%=rs("SERIES_NAME")%></td>
	<td><%=rs("FAMILY_NAME")%></td>
	<%
			rs.movenext
			next 
		end if 
	%>
	
	</tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->