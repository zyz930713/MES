<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
ReportType=request("ReportType")
fromdate=request("fromdate")
fromtime=request("fromtime")
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

if(ReportType="0") then
	sqlstr=GetFinishGoodsSQL(fromdate+" "+fromtime,todate+" "+totime)
end if 

if(ReportType="1") then
	sqlstr=GetFinishGoodsScrapSQL(fromdate+" "+fromtime,todate+" "+totime)
end if 

if(ReportType="2") then
	sqlstr=GetPiecePartsSQL(fromdate+" "+fromtime,todate+" "+totime)
end if 

if(ReportType<>"" and sqlstr<>"") then
	session("SQL")=sqlstr
	'response.write sqlstr
 	rs.open sqlstr,conn,1,3
end if 

%>
<%
	function GetFinishGoodsSQL(starttime,endtime)
		 sql="select d.family_name, c.SERIES_name,a.part_number_tag, sum(a.store_quantity) as Good_Qty, b.item_cost,sum(a.store_quantity*b.item_cost) as Good_Cost "
		 sql=sql+" from job_master_store a , product_model b, SERIES_NEW c, family d"
		 sql=sql+" where  a.part_number_tag= b.item_name and b.series_group_id=c.nid and b.family_id=d.nid(+) "
		 sql=sql+" and line_name<>'ANNEAL' and line_name<>'SECOP' and line_name<>'SUB1' and line_name<>'SUB2' and line_name<>'SUB3' and line_name<>'SUB4' "
		 
		 if(starttime<>"") then
			 sql=sql+" and a.store_time>to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 if(endtime<>"") then
			 sql=sql+" and a.store_time<to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 
		 sql=sql+" group by d.family_name,a.part_number_tag, c.SERIES_NAME,b.item_cost "
		 sql=sql+" order by d.family_name, a.part_number_tag, c.SERIES_NAME,b.item_cost"
		 GetFinishGoodsSQL=sql
		
	end function 
	
	
	function GetFinishGoodsScrapSQL(starttime,endtime)
		sql=" select d.family_name,substr(A.scrap_reference,instr(a.scrap_reference,'-',1,2)+1,instr(a.scrap_reference,'-',1,3)-instr(a.scrap_reference,'-',1,2)-1) as SERIES_NAME, "
		sql=sql+" part_number_tag as modelname, sum(a.scrap_quantity) as scrap_quantity , b.item_cost,sum(a.scrap_quantity*b.item_cost) as scrap_Cost "
		sql=sql+" from job_master_scrap a , product_model b, family d ,series_new c"
		sql=sql+" where  a.part_number_tag= b.item_name and c.family_id=d.nid(+) "
		sql=sql+" and upper(substr(A.scrap_reference,instr(a.scrap_reference,'-',1,2)+1,instr(a.scrap_reference,'-',1,3)-instr(a.scrap_reference,'-',1,2)-1)) =upper(c.series_name) "
		if(starttime<>"") then
			 sql=sql+" and a.scrap_time>to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 if(endtime<>"") then
			 sql=sql+" and a.scrap_time<=to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 sql=sql+" and a.scrap_account='1102'" 
		 sql=sql+" group by d.family_name,substr(A.scrap_reference,instr(a.scrap_reference,'-',1,2)+1,instr(a.scrap_reference,'-',1,3)-instr(a.scrap_reference,'-',1,2)-1), part_number_tag,b.item_cost "
		 sql=sql+" order by d.family_name, substr(A.scrap_reference,instr(a.scrap_reference,'-',1,2)+1,instr(a.scrap_reference,'-',1,3)-instr(a.scrap_reference,'-',1,2)-1),part_number_tag,b.item_cost"

		 GetFinishGoodsScrapSQL=sql
	end function

	function GetPiecePartsSQL(starttime,endtime)
		sql=" select d.family_name,a.SERRIES_NAME,a.item as model_name,sum(a.primary_quantity) as scrap_quantity, unit_cost, sum(txn_value)as total_cost"
		sql= sql+ " from piece_parts_scrap_info a , series_new b, family d  where 1=1 "
		sql=sql+" and  upper(a.SERRIES_NAME)= upper(b.series_name) and b.family_id=d.nid(+) "
		if(starttime<>"") then
			 sql=sql+" and a.txn_date>=to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 if(endtime<>"") then
			 sql=sql+" and a.txn_date<=to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 
		sql=sql+" and (subinv='OOW-KE' or subinv='RAW-KE' or subinv='C10 ASSY' or subinv='FGI-KE') and gl_acct='66.11.71.18.543100.000.000' "
		sql=sql+" group by d.family_name,a.SERRIES_NAME, a.item, unit_cost "
		GetPiecePartsSQL=sql
		response.write SQL
	end function
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
		form1.action="FinancialReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="FinancialReport.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Financial Yield Report</td>
  </tr>
  <tr>
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
	   		 <option value="05:00:00" <% if fromtime="05:00:00" then response.write "Selected" end if%>>05:00:00</option> 
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
			 <option value="05:00:00" <% if fromtime="05:00:00" then response.write "Selected" end if%>>05:00:00</option>
	   			 
  		</select>  
	</td>
	<Td>Report Type</Td>
	<Td><select name="ReportType" id="ReportType">
    		<option  value="0" <% if reporttype="0" then response.write "selected" end if %>>Finish Good</option>
			<option  value="1" <% if reporttype="1" then response.write "selected" end if %>>Finish Scrap</option>
			<option  value="2" <% if reporttype="2" then response.write "selected" end if %>>Piece Parts Scrap</option>
 		 </select>  
	</Td>
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
   			 <td height="20" ><div align="center"><% if isnull(rs(i).value) then response.write "0" else response.write rs(i).value end if %> </div></td>
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
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->