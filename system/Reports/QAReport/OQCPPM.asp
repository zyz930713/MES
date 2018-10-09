 
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

family=request("family")
series=request("series")
subseries=request("subseries")
batchno=request("batchno")

model=request("model")
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
		SQL="select  FA.FAMILY_NAME "
		SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,MODELNAME as model_number,BatchNOlist AS BatchNO"
		SQL=SQL+",j.STARTQTY "	
		SQL=SQL+" ,AQL"
		SQL=SQL+",(SELECT decode(SAMPLEQTY,'100%',j.startqty,SAMPLEQTY) FROM QA_AQL WHERE QTYFROM<=j.STARTQTY AND QTYTO>=j.STARTQTY AND AQL=j.AQL) AS SampleQty"
		SQL=SQL+",(SELECT DECODE(sum(oqcreject),NULL,0,SUM(oqcreject)) FROM JOB_OQC_DETAIL WHERE oqcnid=j.oqcnid) AS RejectQty"
		SQL=SQL+",oqcstarttime AS StartTime,oqcendtime AS EndTime,STARTOPERATOR AS OPERATOR,J.OQCNID "
		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, JOB_OQC j "
		SQL=SQL+" where j.MODELNAME=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID  "
		if(Family<>"") then
			SQL=SQL+" and PM.FAMILY_ID='"+Family+"'"
		end if 
		if(series<>"") then
			SQL=SQL+" and PM.SERIES_GROUP_ID='"+series+"'"
		end if 
		if(subseries<>"") then
			SQL=SQL+" and PM.SERIES_ID='"+subseries+"'"
		end if 
		if(model<>"") then
			SQL=SQL+" and MODELNAME='"+model+"'"
		end if 
		if(batchno<>"") then
			SQL=SQL+" and BatchNOlist like '%"+batchno+"%'"
		end if 
		
		SQL=SQL+" and j.OQCSTARTTIME>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and j.OQCSTARTTIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL +" ORDER BY FAMILY_NAME "
		SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,MODELNAME,BatchNOlist"
	pagename="/reports/newreport/yieldreport.asp"
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
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>
	 
	function GenerateReport()
	{
		form1.action="OQCPPM.asp?Action=GenereateReport"
		form1.submit();
	}
	function LoadNextItem()
	{
		form1.action="OQCPPM.asp?Action=LoadNextItem"
		form1.submit();
	}	
	 
</script>

</head>

<body onLoad="language(<%=session("language")%>);">
<form action="OQCPPM.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
   <tr align="center">
   	<td height="20" width="80"><span id="inner_SearchFamily"></span></td>
     <td height="20" width="100">
	 	<select name="family" id="family" ONCHANGE="LoadNextItem()" style="width:100px">
    	<option value=""></option>
    	<%= getFamily("OPTION",family,factorywhereoutside,"","") %>
 		 </select>  
	</td> 
    <td width="80"><span id="inner_Series"></span></td>
     <td height="20" width="80"><select name="series" id="series" ONCHANGE="LoadNextItem()" style="width:100px">
    	<option value=""></option>
   		 <%= getSeries("OPTION",Series,wherestr1,"","") %>
 	 </select>  </td>
  
    <td width="100"><span id="inner_SubSeries"></span></td>
    <td height="20" width="80">
		<select name="subseries" id="subseries" ONCHANGE="LoadNextItem()" style="width:100px">
    		<option value=""></option>
	   		 <%= getSubSeries("OPTION",SubSeries,wherestr2,"","") %>
 	 	</select>  
	 </td>
   <Td width="80"><span id="inner_SearchPartNumber"></span> </Td>
  	 <td height="20" width="80" colspan="2" align="left">
	 	<select name="model" id="model" style="width:130px">
   			 <option value=""></option>
	   			 <%= getModel("OPTION",model,wherestr3,"","") %>
  		</select>  
	  </td>	   
  </tr>
  <tr align="center">
    <td width="100"><span id="inner_OQCStartTime"></span> </td> 
	 <td colspan="5"><span id="inner_SearchFrom"></span>&nbsp; :
	 	<input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">	
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
	&nbsp;<span id="inner_SearchTo"></span>:
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
		</td>
 
   <Td width="60"><span id="inner_BatchNo"></span></Td>
  	 <td height="20" width="80">
	 	<input name="Batchno" type="text" id="Batchno" value="<%=Batchno%>" size="16">
	  </td>
	   	<Td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  </tr> 
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span id="inner_BrowseData"></span></td>        
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right">
      <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('../NewReport/Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div>
	  </td>
    </tr>
  </table></td>
</tr>
<%if(rs.State > 0 ) then%>
  <tr align="center">
  	<TD height="20" class="t-t-Borrow"><span id="inner_NO"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_Family"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_Series"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_SubSeries"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_Model"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_BatchNo"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_StartQty"></span></TD>
	 <TD class="t-t-Borrow">AQL</TD>
	 <TD class="t-t-Borrow"><span id="td_SampleQty"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_RejectQty"></span></TD>
	  <TD class="t-t-Borrow"><span id="td_OQCStartTime"></span></TD>
	   <TD class="t-t-Borrow"><span id="td_OQCEndTime"></span></TD>
	    <TD class="t-t-Borrow"><span id="td_OpCode"></span></TD>
	 <TD class="t-t-Borrow"><span id="td_DefectInfo"></span></TD>
</tr>
	<%for j=0 to rs.recordcount-1%>
	<tr align="center">
		<TD height="20"><%=(j+1)%></TD>
		 <TD><%=rs(0)%></TD>
		 <TD><%=rs(1)%></TD>
		 <TD><%=rs(2)%></TD>
		 <TD><%=rs(3)%></TD>
		 <TD><%=rs(4)%></TD>
		 <TD><%=rs(5)%></TD>
		 <TD><%=rs(6)%></TD>
		 <TD><%=rs(7)%></TD>
		 <TD><%=rs(8)%></TD>
		 <TD><%=rs(9)%></TD>
		 <TD><%=rs(10)%></TD>
		 <TD><%=rs(11)%></TD>
		 <TD>
		  <%if rs("rejectqty") <> "0" then%>
		  	<input type="button" name="btnDefectInfo" value="Defect Info" onClick="window.open('OQCRejectDetail.asp?oqc_nid=<%=rs("OQCNID")%>')">
		  <%end if%>&nbsp;
		  </TD>
  	<%rs.movenext
	next 
	%>
		</tr>
<%rs.close
end if
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->