<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->


<%
	TransactionType=request.Form("selType")
	if TransactionType="" then
		TransactionType="1"
	end if
	JobNumber=request.Form("txtJobNumber")	
	isPrint=request.form("txtisPrint")
	if isPrint="" then
		isPrint="0"
	end if
	fromdate=request("fromdate")
	todate=request("todate")
	action=request("Action")
	if isnull(fromdate) or fromdate=""  then
		fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
	end if	
	if isnull(todate) or todate=""  then
		todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
	end if
'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
	SQL="select os.job_number,os.batchno,os.qty,os.subjoblist,os.actualqty,os.operator,os.intimestamp,os.transactiontype, "
	SQL=SQL+" (select defect_name||'('||defect_chinese_name||')-'||station_id from defectcode where nid=os.remark) as remark,"
	SQL=SQL+"os.isprint,os.new_jobnumber,jm.part_number_tag from offline_store os,job_master jm where os.job_number = jm.job_number "	
	SQL=SQL+" and os.intimestamp>to_date('"& fromdate & " 00:00:00','yyyy-mm-dd hh24:mi:ss')"
	SQL=SQL+" and os.intimestamp<=to_date('"& todate & " 23:59:59','yyyy-mm-dd hh24:mi:ss')"
	
	if TransactionType <> "" and TransactionType <> "-2"  then
		SQL=SQL+" and os.transactiontype='"&TRANSACTIONTYPE&"'"
	end if
	if isPrint<>"2" then
		SQL=SQL+" and  os.ISpRINT='"&isPrint&"'"
	end if 
	if JobNumber<>"" then
		SQL=SQL+" and  os.job_number='"&JobNumber&"'"
	end if
	SQL=SQL+" order by os.job_number,os.transactiontype"	
	session("sql")=sql
	rs.open SQL,conn,1,3 
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>

<script type="text/javascript">
	function QueryData()
	{
		form1.action="DefectHistory.asp?Action=GenereateReport"
		form1.submit();
	}
	function ExportExcel()
	{
		window.open("ExportExcel.asp");
	}
	function pageLoad()
	{	
		<%if action="GenereateReport" then %>
			document.all.btnExport.disabled=false;
		<%else%>
			document.all.btnExport.disabled=true;
		<%end if%>
	}
</script>
</head>
<body bgcolor="#339966" onLoad="pageLoad()" >
<table width="95%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<form id="form1" method="post" name="form1">
	<tr>
		<Td height="20" colspan="8" class="t-t-DarkBlue"  align="center">Defect History 缺陷历史</Td>
	</tr>	
	<tr><td>
		<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
		<tr>
		<Td>Job Number 工单号</Td>
		<Td  ><input name="txtJobNumber" type="text" id="txtJobNumber" value=<%=JobNumber%>></Td>
		<td>Transaction Type 类型</td>
		<td ><select id="selType"  name="selType" >
			<option value="-2" <%if TransactionType="-2" then%> selected <% end if%>>--All--</option>
<!--			<option value="-1" <%if TransactionType="-1" then%> selected <% end if%>>Normal</option>-->
<!--			<option value="0" <%if TransactionType="0" then%> selected <% end if%>>None</option>-->
			<option value="1" <%if TransactionType="1" then%> selected <% end if%> >Rework</option>
			<option value="2" <%if TransactionType="2" then%> selected <% end if%>>Scrap</option>
<!--			<option value="3" <%if TransactionType="3" then%> selected <% end if%>>Readjust</option>-->
			<option value="4" <%if TransactionType="4" then%> selected <% end if%>>Retest</option>
<!--			<option value="5" <%if TransactionType="5" then%> selected <% end if%>>Slapping</option>-->
		</select>	</td>

		<Td >Job Release 是否已开工单</Td>
		<Td ><select id="txtisPrint"  name="txtisPrint" >
			<option value="2" <%if isPrint="2" or  isPrint="" then%> selected <% end if%>>--All--</option>
			<option value="0" <%if isPrint="0" then%> selected <% end if%>>No</option>
			<option value="1" <%if isPrint="1" then%> selected <% end if%>>Yes</option>			
			</select>
 		 <td >DateTime  发生时间</td> 
		 <td >From 从:</td>
		 <td>
		 <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
		
			<script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.fromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
			</script>
			<!--			
			<select name="fromtime" id="fromtime">
			  <option value="08:10:00" <% if fromtime="08:10:00" then response.write "Selected" end if%>>08:10:00</option>
			  <option value="12:00:00" <% if fromtime="12:00:00" then response.write "Selected" end if%>>12:00:00</option>
              <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="15:00:00" <% if fromtime="15:00:00" then response.write "Selected" end if%>>15:00:00</option>
			  <option value="21:00:00" <% if fromtime="21:00:00" then response.write "Selected" end if%>>21:00:00</option>
              <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
            </select>
			-->
			</td>
			<td >To 到:</td>
			<td>
			<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
			<script language=JavaScript type=text/javascript>
				function calendar2Callback(date, month, year)
				{
				document.all.todate.value=year + '-' + month + '-' + date
				}
				calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
			</script>
			&nbsp;
			<!--
	  		<select name="totime" id="totime">
             <option value="08:10:00" <% if fromtime="08:10:00" then response.write "Selected" end if%>>08:10:00</option>
			  <option value="12:00:00" <% if fromtime="12:00:00" then response.write "Selected" end if%>>12:00:00</option>
              <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="15:00:00" <% if fromtime="15:00:00" then response.write "Selected" end if%>>15:00:00</option>
			  <option value="21:00:00" <% if fromtime="21:00:00" then response.write "Selected" end if%>>21:00:00</option>
              <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
          </select>
		  -->

	  <Td align="center">
	  		<input name="btnQuery" type="button"  id="btnQuery" onClick="QueryData()" value="Query 查询">
			<input name="btnExport" type="button"  id="btnExport" onClick="ExportExcel()" value="导出Excel">
	  </Td>
	</tr>
	</table></td></tr>
	<tr>
		<Td colspan="8">
			<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
				<Tr class="today" align="center" >
					<Td>Job Number <br>工单号</Td>					
					<Td>Model Name<br> 型号 </Td>
					<Td>Sub Job List<br>子工单号</Td>
					<Td>Qty <br>数量</Td>
					<Td>Actual Qty<br> 实际数量</Td>
					<Td>Operator <br>工号</Td>
					<Td>DateTime <br>发生时间</Td>
					<Td>Remark 备注</Td>
					<Td>Job Release<br> 是否已开工单</Td>
					<Td>New Job Number<br> 打印新工单号</Td>
					<td>Transaction Type<br> 类型</td>
			  </Tr>
				<%if request.querystring("Action")="GenereateReport" then
					if rs.eof then 
				%>
					<tr><td align="center" colspan="12">No Records 没有记录</td></tr>
				<%
					end if
					while not rs.eof
				%>
				<Tr align="center">					
					<Td ><%=rs("JOB_NUMBER")%></Td>
					<Td><%=rs("part_number_tag")%></Td>
					<Td><%=rs("SUBJOBLIST")%>&nbsp;</Td>
					<Td><%=rs("QTY")&"&nbsp;"%></Td>
					<Td><%=rs("ACTUALQTY")&"&nbsp;"%></Td>
					<Td><%=rs("OPERATOR")&"&nbsp;"%></Td>
					<Td><%=rs("INTIMESTAMP")&"&nbsp;"%></Td>
					<Td><%=rs("REMARK")&"&nbsp;"%></Td>
					<Td><%if rs("ISPRINT")="1" then response.write "Yes" end if 
						  if rs("ISPRINT")="0" then response.write "NO" end if%>&nbsp;					
					</Td>
					<Td><%=rs("NEW_JOBNUMBER")&"&nbsp;"%></Td>					
					<Td>
						  <%if rs("TransactionType")="-1" then response.write "Normal" end if %> 
						  <%if rs("TransactionType")="0"  then response.write "None" end if  %>  
						  <%if rs("TransactionType")="1" then response.write "Rework" end if %>  
						  <%if rs("TransactionType")="2" then response.write "Scrap" end if %>  
						  <%if rs("TransactionType")="3" then response.write "Readjust" end if %> 
						  <%if rs("TransactionType")="4" then response.write "Retest" end if %> 
						  <%if rs("TransactionType")="5" then response.write "Slapping" end if %> 
					&nbsp;</Td>
				</Tr>
				<% 	rs.movenext
					wend 
					
				end if 
				%>
			</table>
		</Td>
	</tr>
</form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
