<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
JobNumber=request.Form("txtJobNumber")
action=request.querystring("Action")
if action="GenereateReport" then
		SQL="select  * from label_print_history where job_number='"+JobNumber+"' "
		rs.open SQL,conn,1,3
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>

<script>
		 
		
		function QueryData()
		{
			 
			form1.action="LabelPrintHistory.asp?Action=GenereateReport"
			form1.submit();
		}
	
</script>
<link href="../CSS/CMMS.css" rel="stylesheet" type="text/css">
</head>
<body style=" font-size:12px; font-family:Arial, Helvetica, sans-serif">
<table width="90%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<form id="form1" method="post" name="form1">
	<tr>
		<Td height="20" colspan="8" class="t-t-DarkBlue"  align="center">LabelPrint History data</Td>
	</tr>
	<tr>
    
		<Td>工单号(Job Number)</Td>
		<Td  ><input name="txtJobNumber" type="text" id="txtJobNumber" value=<%=JobNumber%>></Td>
		<Td>&nbsp;</Td>
		<Td>&nbsp;</Td>
	 	<Td>&nbsp;</Td>
		<Td>&nbsp;</Td>
	</tr>
	 
	<tr>
	  <Td colspan="8" align="center"><input name="btnQuery" type="button" class="t-c-greenCopy" id="btnQuery" onClick="QueryData()" value="Query (查询)">
				</Td>
	</tr>
	<tr>
		<Td colspan="8">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
				<Tr class="today" >
					<Td>
						Job Number(工单号)					</Td>
					<Td>
						Batch Number(入库单号)					</Td>
					<Td>
						Sub Job List(子工单号)					</Td>
					<Td>
						Qty(数量)					</Td>
					<Td>
						Actual Qty(实际入库数量)					</Td>
		
					<Td>
						DateTime(入库时间)					</Td>
					<td>
						Transaction Type					</td>
			  </Tr>
				<%
					if request.querystring("Action")="GenereateReport" then
						while not rs.eof
				%>
				<Tr>
		
					<Td >
						<%=rs("JOB_NUMBER")%>					</Td>
					<Td>
						<%=rs("BATCHNO")%>					</Td>
					<Td>
						<%=rs("SUBJOBLIST")&"&nbsp;"%>					</Td>
					<Td>
						<%=rs("QTY")&"&nbsp;"%>					</Td>
					<Td>
						<%=rs("ACTUALQTY")&"&nbsp;"%>					</Td>
				
					<Td>
						<%=rs("printtime")&"&nbsp;"%>					</Td>
					
					<Td>
						  <%if rs("TransactionType")="-1" then response.write "Normal" end if %> 
						  <%if rs("TransactionType")="0"  then response.write "None" end if  %>  
						  <%if rs("TransactionType")="1" then response.write "Rework" end if %>  
						  <%if rs("TransactionType")="2" then response.write "Scrap" end if %>  
						  <%if rs("TransactionType")="3" then response.write "Readjust" end if %> 
						  <%if rs("TransactionType")="4" then response.write "Change Model" end if %> 
						  <%if rs("TransactionType")="5" then response.write "Slapping" end if %> 
&nbsp;					</Td>
				</Tr>
				<%  
						rs.movenext
						wend 
					end if 
				%>
			</table>		</Td>
	</tr>
</form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
