<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->

<%
action=request("Action")

if action="1" then
		SQL="select  FA.FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME "

		SQL=SQL+", SUM(j.JOB_START_QUANTITY) AS INPUT, SUM(j.JOB_GOOD_QUANTITY) AS OUTPUT , "	
		SQL=SQL+"  to_char(round((SUM(j.JOB_GOOD_QUANTITY) / sum(j.JOB_START_QUANTITY)),4)*100) || '%' as FPY "
		SQL=SQL+" ,MAX(SS.TARGET_FIRSTYIELD) as target_FPY,"
		
		SQL=SQL+" SUM((select SUM(decode(nvl2(translate(action_value,'\1234567890','\'),'0','1'),0,0,1,action_value)) FROM JOB_ACTIONS JA "
  		SQL=SQL+" WHERE JA.JOB_NUMBER=J.JOB_NUMBER AND JA.SHEET_NUMBER=J.SHEET_NUMBER AND JA.ACTION_ID='AC00000550')) AS SLAPPING "
		
		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, job  J "
		SQL=SQL+" where j.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID "
	
		'SQL=SQL+" and j.JOB_NUMBER=JA.JOB_NUMBER(+)"
		'SQL=SQL+" and j.SHEET_NUMBER=JA.SHEET_NUMBER(+)"
		'SQL=SQL+" and ja.action_id='AC00000550' "
		SQL=SQL+" and j.close_time>to_date(to_char(sysdate-1,'yyyy-mm-dd') || ' 14:30:00','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and j.close_time<=to_date(to_char(sysdate,'yyyy-mm-dd') || ' 14:30:00','yyyy-mm-dd hh24:mi:ss')"
		
		SQL=SQL+" GROUP BY FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME"
		
		SQL=SQL +" ORDER BY FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME "
	
	session("SQL")=SQL
	'response.write sql
	'response.end
	rs.open SQL,conn,1,3
	Report_title="FPY Report"
	
end if

if action="2" then
		SQL="select  FA.FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME "

		SQL=SQL+", SUM(jms.INPUT_QUANTITY) as Input,SUM(jms.STORE_QUANTITY) AS OutPut, "	
		SQL=SQL+" to_char(round((SUM(jms.STORE_QUANTITY) / DECODE(SUM(jms.INPUT_QUANTITY),0,1,SUM(jms.INPUT_QUANTITY))),4)*100) || '%' as FPY"
		SQL=SQL+" ,MAX(SS.TARGET_FIRSTYIELD) as target_FPY"
		
		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM,JOB_MASTER_STORE JMS "
		SQL=SQL+" where JMS.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID  AND ( JMS.STORE_TYPE='N' OR JMS.STORE_TYPE='R') "
		
		SQL=SQL+" and JMS.STORE_TIME>to_date(to_char(sysdate-1,'yyyy-mm-dd') || ' 14:30:00','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and JMS.STORE_TIME<=to_date(to_char(sysdate,'yyyy-mm-dd') || ' 14:30:00','yyyy-mm-dd hh24:mi:ss')"
		
		SQL=SQL+" GROUP BY FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME"
		
		SQL=SQL +" ORDER BY FAMILY_NAME,SERIES_NAME,SUBSERIES_NAME "
	
	session("SQL")=SQL
	'response.write sql
	'response.end
	Report_title="Final Yield Report"
	rs.open SQL,conn,1,3
end if


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script>
	function GenerateReport(flag)
	{
		form1.action="AutoReport.asp?Action="+flag
		form1.submit();
	}
</script>

</head>

<body>
<form action="AutoReport.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Auto Report</td>
  </tr>
  <tr>
	<Td><input type="button" value="FPY" id="btnAutoReportFPY" name="btnAutoReportFPY" onclick="GenerateReport('1')"></Td>
	<Td><input type="button" value="Final Yield" id="btnAutoReportFinal" name="btnAutoReportFinal" onclick="GenerateReport('2')"></Td>
	<Td colspan="6">&nbsp;</Td>
	 
  </Tr>
</table>
</form>
 
 <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  
  <tr>
  	<td colspan="18"><%=Report_title%></td>
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
			 		<%if(isnull(rs(i).value) or rs(i).value="") then
							response.write "0"
						else
							response.write rs(i).value
						end if%> </div></td>
			 
			 <%
			 	if(rs.Fields(i).name="INPUT") then
					input=input+cdbl(rs(i).value)
				end if
				
				if(rs.Fields(i).name="OUTPUT") then
					moutput=moutput+cdbl(rs(i).value)
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
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->