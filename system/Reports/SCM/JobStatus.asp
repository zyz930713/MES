<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
 

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request("jobnumber")

modelname=request("modelname")
planer=request("planer")
fromdate=request("fromdate")
fromtime=request("fromtime")
action=request("Action")

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
		SQL="select Job_Number,part_number_tag as Model_Name,start_quantity as Input_Qty,create_code as Planer from job_master where store_status='0' "
		if(jobnumber<>"") then
			SQL=SQL+" AND job_number='"+jobnumber+"'"
		end if
		 
		 if(modelname<>"") then
			SQL=SQL+" AND part_number_tag='"+modelname+"'"
		end if
		
		
		if(planer<>"") then
			SQL=SQL+" AND create_code='"+planer+"'"
		end if
		
		SQL=SQL+" and input_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and input_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		
 
	'pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
	
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
<script>
	function LoadNextItem()
	{
		form1.action="JobStatus.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		form1.action="JobStatus.asp?Action=GenereateReport"
		form1.submit();
	}
	
	 
</script>

</head>

<body>
<form action="JobStatus.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Job Status Report</td>
  </tr>
  <tr>
    <td height="20">Job Number </td>
     <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" size="16"></td>
  
  
    <td>Model Name</td>
     <td height="20"><input name="modelname" type="text" id="modelname" value="<%=modelname%>" size="16"></td>
  
    <td>Planer</td>
    <td height="20"><input name="planer" type="text" id="planer" value="<%=planer%>" size="16"></td>
   <Td>&nbsp;</Td>
  	 <td height="20">
	  &nbsp;
	  </td>
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
  	<Td colspan="3"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('../NewReport/Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="10"> </td>
  </tr>
  <tr>
	 
    <td height="20" class="t-t-Borrow"><div align="center">Job_Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Model_Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Input_Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">First Out Put</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">FPY</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Slapping Input</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Slapping Out</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Rework Input</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Rework Output</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Job Final Out Put</div></td>
</tr>
	<%  if(rs.State > 0 ) then	
			for i=0 to rs.recordcount-1
	%>
	<tr>
	 	<td height="20" ><div align="center"><%=rs(0).value%> &nbsp;</div></td>
		<td height="20" ><div align="center"><%=rs(1).value%> &nbsp;</div></td>
		<td height="20" ><div align="center"><%=rs(2).value%> &nbsp;</div></td>
		<%
			TotalInput=rs(2).value
			'get first past yield
			SQL="select sum(job_good_quantity) from job where job_number='"+rs(0).value+"'"
			set rsJobInfo=server.createobject("adodb.recordset")
			rsJobInfo.open SQL,conn,1,3
			FirstOutPut=0
			
			if rsJobInfo.recordcount>0 then
				FirstOutPut=rsJobInfo(0).value
			else
				FirstOutPut=0
			end if
		%>
		<td height="20" ><div align="center"> <%=FirstOutPut%> &nbsp;</div></td>
		<%
			FPY=0
			if cdbl(TotalInput)>0 then
				FPY=cstr(round(cdbl(FirstOutPut)/cdbl(TotalInput)*100 ,2))+"%"
			end if
		%>
		
		<td height="20" ><div align="center"> <%=FPY%> &nbsp;</div></td>
		
		<%
			sql="select sum(ja.action_value) from job_actions ja ,action a where  ja.action_id=a.nid and  job_number='"+rs(0).value+"'"
			sql=sql+" and a.mother_action_id='AN00000103'"
			set rsJobTotalSlapping=server.createobject("adodb.recordset")
	 
			rsJobTotalSlapping.open sql,conn,1,3
			TotalSlappingQty=0
			if (rsJobTotalSlapping.recordcount>0)then
				TotalSlappingQty=rsJobTotalSlapping(0).value
			end if
			
			sql="select ja.action_value from job_actions ja ,action a where  ja.action_id=a.nid and  job_number='"+rs(0).value+"'"
			sql=sql+" and a.mother_action_id='AN00000141'"
			set rsJobGoodSlapping=server.createobject("adodb.recordset")
			rsJobGoodSlapping.open sql,conn,1,3
			SlappingGoodQty=0
			if (rsJobGoodSlapping.recordcount>0)then
				SlappingGoodQty=rsJobGoodSlapping(0).value
			end if
			
		%>
		<td height="20" ><div align="center"><%=TotalSlappingQty%>&nbsp;</div></td>
		<td height="20" ><div align="center"><%=SlappingGoodQty%>  &nbsp;</div></td>
		
		<%
			'Get Retest reject Qty
			SQL1="select  decode(SUM(a.SAMPLEQTY),null,0,SUM(a.SAMPLEQTY)),decode(SUM(A.REJECTQTY),null,0,SUM(A.REJECTQTY)) from Job_ReTest_detail   a, "
			SQL1=SQL1+"(select batchno,max(testsequence) as testsequence from Job_ReTest_detail "
			SQL1=SQL1+"group by batchno)b,"
			SQL1=SQL1+" Job_ReTest C "
			SQL1=SQL1+" where a.batchno=b.batchno and a.testsequence=b.testsequence"
			SQL1=SQL1+" AND A.BATCHNO=C.BATCHNO AND C.isreleasetooqc='1' AND c.BATCHNO LIKE '"+rs(0).value+"%'"
			
			set rsRetest=server.createobject("adodb.recordset")
			rsRetest.open SQL1,conn,1,3
			RetestTotalIn=0
			RetestTotalReject=0
			
			if rsRetest.recordcount>0 then
				RetestTotalIn=rsRetest(0)
				RetestTotalReject=rsRetest(1)
			end if
			
		%>
		<td height="20" ><div align="center"><%=RetestTotalIn%> &nbsp;</div></td>
		<td height="20" ><div align="center"><%=cdbl(RetestTotalIn)-cdbl(RetestTotalReject)%>  &nbsp;</div></td>
		<td height="20" ><div align="center">  &nbsp;</div></td>
  	<%
			rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
   <tr>
   
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
     <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
     <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
