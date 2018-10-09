<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<%
fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))

jobnumber=request("txtjobnumber")
sheetnumber=request("txtSubjobnumber")
line=request("dropLineName")


partnumber=request("txtpartnumber")
fromdate=request("fromdate")
todate=request("todate")
time0=now   
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

todate=request("todate")
if isnull(todate) or todate=""  then
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
	
if(request.QueryString("action")="1")then
	sql="select * from job a where 1=1"
	if jobnumber<>"" then
		sql=sql+" and job_number='"+jobnumber+"'"
	end if 
	if sheetnumber<>"" then
		sql=sql+" and sheet_number='"+sheetnumber+"'"
	end if 
	if line<>"" then
		sql=sql+" and line_name='"+line+"'"
	end if 
	
	if partnumber<>"" then
		sql=sql+" and part_number_tag='"+partnumber+"'"
	end if 
	
	
	if fromdate<>"" then
		sql=sql+" and close_time>=to_date('"+fromdate+"','yyyy-mm-dd hh24:mi:ss')"
	end if 
	if todate<>"" then
		sql=sql+" and close_time<=to_date('"+todate+" 23:59:59','yyyy-mm-dd hh24:mi:ss')"
	end if 
	sql=sql+" and not exists(select 1 from job_master_store_pre where instr(sub_job_numbers,get_sub_job_number(a.job_number,a.sheet_number))>0 ) "
	sql=sql+" order by job_number,sheet_number"
	
	session("SQL-SubAuto")=sql
	rs.open sql,conn,1,3
end if 

			
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	function searchdata()
	{
		if(document.getElementById("dropLineName").value=="")
		{
			alert('Please select one line!\n请选择线别!');
			return;
		}
		if(document.getElementById("fromdate").value=="" )
		{
			alert('Please select from date!\n请输入开始时间!');
			return;
		}
		if(document.getElementById("todate").value=="" )
		{
			alert('Please select to date!\n请输入结束时间!');
			return;
		}
		form1.action="pickjob.asp?action=1";
		form1.submit();
	}
	
	function pickdata()
	{
		var count=document.getElementById("count").value;
		var selectCount=0;
		var mainJob="";
		var jobNumber="";
		var sheetNumber="";
		
		for(var i=0;i<=count;i++)
		{
			if(document.getElementById("selectck"+i).checked)
			{
				jobNumber=document.getElementById("jobnumber"+i).value;	
				sheetNumber=document.getElementById("sheetnumber"+i).value			
				if(mainJob==""){
					mainJob=jobNumber;
				}else if(mainJob != jobNumber){
					alert("Cannot mix job number.\n不能混工单.");
					break;
				}				
				
				selectCount=selectCount+1;								
				document.getElementById("txtSelectJobNumber").value=document.getElementById("txtSelectJobNumber").value+jobNumber+",";
				document.getElementById("txtSelectSheetNumber").value=document.getElementById("txtSelectSheetNumber").value+sheetNumber+",";	
			}
		}
		if(selectCount>15)
		{
			alert('Qty of selected job has exceed 15.\n选择的工单数量超过了15个.');
			return;
		}

		if(document.getElementById("txtSelectJobNumber").value=="")
		{
			alert('Please one record!\n请选择至少一行数据!');
			return;
		}
		form1.action="Store1.asp?store_type=subStore";
		form1.submit();
	}
	
	function SelectAll()
	{
		var count=document.getElementById("count").value;
		var selAll=document.getElementById("sel_all").checked;
		for(var i=0;i<=count;i++)
		{
			document.getElementById("selectck"+i).checked=selAll;
		}
	}
	
	function CancelAll()
	{
		var count=document.getElementById("count").value;
		for(var i=0;i<=count;i++)
		{
			document.getElementById("selectck"+i).checked=false;
		}
	}
	function goPrint()
	{
		location.href="SelectFactory.asp";
	}
	function goPrint2()
	{
		location.href="../Scrap/SelectFactory.asp";
	}
	
</script>
</head>

<body bgcolor="#339966">
<span id="erroralarm"></span>
<form  method="post" name="form1" target="_self" >
<table width="98%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-t-DarkBlue"  align="center" >HFL Store 入库</div></td>
  </tr>
   <tr>
    <td height="20" ><div align="left">Job Number工单号</div></td>
	<td height="20"><input name="txtjobnumber" id="txtjobnumber" value="<%=jobnumber%>" type="text"></td>
	 <td height="20">Part Number 型号</td>
	 <td height="20"><input name="txtpartnumber" id="txtpartnumber" value="<%=partnumber%>"></td>   
	 <td height="20"><div align="left">Line 线别</div></td>
	<td height="20">
		<select name="dropLineName" id="dropLineName" >
			<option value=""></option>
    		<%=getLine2("OPTION",line,"where line_name like '85%'","","")%>
 		 </select>  
	</td>	
  </tr>
   <tr>
   <td><div align="left">Close Time 关闭时间</div></td>
    <td colspan="2"><table  border="0" cellpadding="0" cellspacing="0"><tr><td>From 从</td><td>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script>
	</td><td>&nbsp;To 到</td><td>
	<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
	<script language=JavaScript type=text/javascript>
		function calendar2Callback(date, month, year)
		{
		document.all.todate.value=year + '-' + month + '-' + date
		}
		calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
	</script>
	</td></tr></table>
	</td><td colspan="3">
		<input type="button" name="btnSearch" id="btnSearch" value="Query 查询" onclick="searchdata()">
		<input type="button" name="btnpickdata" id="btnpickdata" value="Next 下一步" onclick="pickdata()">
		<!--
		<input type="button" name="btnPrint" id="btnPrint" value="良品入库打印" onclick="goPrint()">
		-->
		<input type="button" name="btnClose" id="btnClose" value="Close 关闭" onclick="javascript:window.close()">
	</td>
  </tr>
   <tr>
	
	<!--
	<td align="right">
		<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('PickJobExportExel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span>
	</td>
	-->		
	</tr>
<%if(request.QueryString("action")<>"")then %>	
  <tr>
  	<Td colspan="6">
		<table  width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr align="center">				
				<td class="t-b-blueReal"><input type="checkbox" name="sel_all" id="sel_all" onClick="SelectAll()"></td>
				<td class="t-b-blueReal">No<br>序号</td>
				<td class="t-b-blueReal" >Job Number<br>工单号</td>
				<td class="t-b-blueReal">Sheet Number<br>子工单号</td>
				<td class="t-b-blueReal">Line<br>线别</td>
				<td class="t-b-blueReal">Part Number<br>型号</td>
				<td class="t-b-blueReal">Start Time<br>开始时间</td>
				<td class="t-b-blueReal">Close Time<br>结束时间</td>
				<td class="t-b-blueReal">Start Qty<br>开始数量</td>
				<td class="t-b-blueReal">Good Qty<br>良品数量</td>
				<td class="t-b-blueReal">Scrap Qty<br>报废数量</td>
			</tr>
			<%i=0
			while  not rs.eof						
			%>
				<tr>
				<td align="center">
				  <input type="checkbox" name="selectck<%=cstr(i)%>" value="selectck"></td>
				  <td><%=i+1%></td>
				<td><%=rs("job_number")%><input type="hidden" name="jobnumber<%=cstr(i)%>" value="<%=rs("job_number")%>"></td>
				<td><%=rs("sheet_number")%><input type="hidden" name="sheetnumber<%=cstr(i)%>" value="<%=rs("sheet_number")%>"></td>
				<td><%=rs("line_name")%></td>
				<td><%=rs("part_number_tag")%></td>
				<td><%=rs("start_time")%></td>
				<td><%=rs("close_time")%></td>
				<td><%=rs("JOB_START_QUANTITY")%></td>
				<td><%=rs("JOB_GOOD_QUANTITY")%></td>
				<td><%=cstr(cint(rs("JOB_START_QUANTITY"))-cint(rs("JOB_GOOD_QUANTITY")))%></td>
				</tr>
			<%i=i+1
			rs.movenext
			wend 
			%>
		</table>
	</Td>
  </tr>
 <%end if %>
 </table>
<input type="hidden" name="txtSelectJobNumber" id="txtSelectJobNumber" >
<input type="hidden" name="txtSelectSheetNumber" id="txtSelectSheetNumber" >
<input type="hidden" name="count" value="<%=cstr(i-1)%>">
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->