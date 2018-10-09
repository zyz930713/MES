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
batchno=request("batchno")
jobnumber =request("jobnumber")
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")

jobtype=request("JobType")
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
wherestr1=" where 1=2"
wherestr2=" where 1=2"
wherestr3=" where 1=2"

if not isnull(family) and family<>"" then
	 if(action="LoadNextItem" OR action="GenereateReport" ) then
	 	wherestr1=" where Family_id='"&family&"'"
	 end if
end if

if not isnull(Series) and Series<>"" then
	if(action="LoadNextItem" OR action="GenereateReport" ) then
	 	wherestr2=" where Series_id='"&Series&"'"
	 end if
end if

if  not isnull(SubSeries) and SubSeries<>"" then
	if(action="LoadNextItem" OR action="GenereateReport") then
	 	wherestr3=" where Series_id='"&SubSeries&"'"
	 end if
end if

'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
		SQL=" select a.batchno,a.retestreceivetime,a.receiveop,a.startqty,b.subjoblist,b.transactiontype,j.line_name "
		SQL=SQL+" from job_master j, Job_ReTest a , label_print_history b where b.batchno=a.batchno AND b.job_number=j.job_number "
		if batchno<>"" then
		 	SQL=SQL+" AND a.BATCHNO='"+batchno+"'"
		end if 
		if(jobnumber<>"") then
			sql=sql+" and b.job_number='"+jobnumber+"'"
		end if 
	    SQL=SQL+" AND a.retestreceivetime>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and a.retestreceivetime<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss') "
		SQL=SQL+" and ((a.batchno like '_______-_-%')or (a.batchno like '_______-__-%')  or (a.batchno like '________-_-%') or (a.batchno like '________-__-%'))"
       	SQL=SQL+" order by  a.batchno,a.retestreceivetime desc"
		
	session("SQL")=SQL
	session("FromDateTime") =fromdate & " " &fromtime
	session("ToDateTime") =todate & " " &totime
	'response.write sql
	'RESPONSE.END
	rs.open SQL,conn,1,3
	
end if

function getMainJobNumber(JobNumber)
		arrJob=split(JobNumber,"-")
		NewJobNumber=""
		 if arrJob(1)="E" or arrJob(1)="R" then
			NewJobNumber=arrJob(0)&"-"&arrJob(1)
		else
			NewJobNumber=arrJob(0)
		end if
		getMainJobNumber=NewJobNumber
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
	function LoadNextItem()
	{
		form1.action="RetestIn.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		 

		form1.action="RetestIn.asp?Action=GenereateReport"
		form1.submit();
	}
	
	function GenerateDetailReport()
	{
		 
		form1.action="RetestIn.asp?Action=GenereateDetialReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="RetestIn.asp" method="post" name="form1" target="_self" id="form1">
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-b-midautumn">Retest In Report</td>
  </tr>
   
  
  <Tr>
  	<td>BatchNo</td>
	<td><input name="batchno" type="text" id="batchno" value="<%=batchno%>" size="16"></td>
	<td>JobNumber</td>
	<td><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" size="16"></td>
	
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
  	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()">
		<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('RetestIN_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span>
	</Td>
  	 
  </Tr>
</table>
</form>
<table width="98%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
 
  <tr>
	<td class="t-t-Borrow">BatchNO</td>
	<td class="t-t-Borrow">子工单号</td>
	<td class="t-t-Borrow">型号</td>
	<td class="t-t-Borrow">接受时间</td>
	<td class="t-t-Borrow">接受人</td>
	<td class="t-t-Borrow">接受数量</td>
	<td class="t-t-Borrow">工单类型</td>
	<td class="t-t-Borrow">线别</td>
  </tr>
  
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	 	<td><%=rs(0).value%></td>
		<td><%=rs(4).value%></td>
 		<%
			SQL="SELECT * FROM label_print_history WHERE BATCHNO='"+rs(0).value+"' AND (SUBJOBLIST LIKE '%70%' OR SUBJOBLIST LIKE '%71%' OR SUBJOBLIST LIKE '%72%')"
			set rsModelType=server.createobject("adodb.recordset")
			rsModelType.open SQL,conn,1,3
			if(rsModelType.recordcount=0) then
				sql="select part_number_tag from job where job_number='"+getMainJobNumber(rs(0).value)+"' and job_type='N' "
				set rsModel1=server.createobject("adodb.recordset")
				rsModel1.open SQL,conn,1,3	
				if(rsModel1.recordcount>0) then
					modelname=rsModel1(0)
				end if 
			else
				sql="select part_number_tag from job where job_number='"+getMainJobNumber(rs(0).value)+"' and job_type='C' "
				set rsModel2=server.createobject("adodb.recordset")
				rsModel2.open SQL,conn,1,3
				if(rsModel2.recordcount>0) then
					modelname=rsModel2(0)
				end if 	
			end if
		%>
		<td><%=modelname%> </td>
		<td><%=rs(1).value%> </td>
		<td><%=rs(2).value%></td>
		<td><%=rs(3).value%></td>
		<td>
		  <%if rs("TransactionType")="-1" then response.write "Normal" end if %> 
		  <%if rs("TransactionType")="0"  then response.write "None" end if  %>  
		  <%if rs("TransactionType")="1" then response.write "Rework" end if %>  
		  <%if rs("TransactionType")="2" then response.write "Scrap" end if %>  
		  <%if rs("TransactionType")="3" then response.write "Readjust" end if %> 
		  <%if rs("TransactionType")="4" then response.write "Change Model" end if %> 
		  <%if rs("TransactionType")="5" then response.write "Slapping" end if %> 
		</td>
		<td><%=rs(6).value%></td>
		<%TotalQty=TotalQty+cint(rs(3).value)%>
		
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
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td  class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">总数量</td>
	<td  class="t-t-Borrow"><%=TotalQty%>&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
  </tr>
   
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->