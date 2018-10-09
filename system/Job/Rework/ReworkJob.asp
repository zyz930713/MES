<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
dim ReworkJobNumber
dim ReworkType
dim StartTime
dim EndTime
dim OperatorCode
dim WhereStr

ReworkJobNumber=request("txtreworkjobnumber")
ReworkType=request("dropReworkType")
StartTime=request("txtfromdate")
EndTime=request("txttodate")
OperatorCode=request("txtOperatorCode")

	if ReworkJobNumber<>"" then
		WhereStr=WhereStr &" AND REWORK_JOB_NUMBER='"& ReworkJobNumber &"'"
	end if 
	if ReworkType<>"" AND ReworkType<>"-1"  then
		WhereStr=WhereStr &" AND REWORK_TYPE='"& ReworkType &"'"
	end if 
	if StartTime<>"" then
			WhereStr=WhereStr &" AND REWORK_START_TIME>=to_date('"& StartTime &" 00:00:00','yyyy-mm-dd hh24:mi:ss')"
		else
			StartTime=dateadd("d",-7,date())
			WhereStr=WhereStr &" AND REWORK_START_TIME>=to_date('"& StartTime &" 00:00:00','yyyy-mm-dd hh24:mi:ss')"
		
	end if 
	if EndTime<>"" then
		WhereStr=WhereStr &" AND REWORK_END_TIME<=to_date('"& EndTime &" 23:59:59','yyyy-mm-dd hh24:mi:ss')"
	end if 
	if OperatorCode<>"" then
		WhereStr=WhereStr &" AND OPERATOR_CODE='"& OperatorCode &"'"
	end if 
	pagename="/Job/Rework/ReworkJob.asp"
	FactoryRight "J."
	
	SQL="select * FROM JOB_REWORK where 1=1" & WhereStr & " ORDER BY REWORK_END_TIME DESC,REWORK_JOB_NUMBER,TO_NUMBER(SUB_JOB_NUMBER)"
	session("SQLQuery")=SQL
	rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
	
<!--#include virtual="/Language/Job/Reworkjob/Lan_Job.asp" -->
</head>

<body onLoad="language(); language_page();language_jobnote()">
<form name="form1" method="get" action="/Job/rework/ReworkJob.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span></td>
    <td height="20"><input name="txtreworkjobnumber" type="text" id="jobnumber" value="<%=ReworkJobNumber%>"></td>
    <td><span id="inner_SearchReworkType"></span></td>
    <td><SELECT id="dropReworkType"  name="dropReworkType" style="width:100px;FONT-SIZE:10pt">	
								<option  <% if ReworkType="-1" or ReworkType="" then response.write "SELECTED" end if%> value="-1">-Select-</option>
								<option <% if ReworkType="0"  then response.write "SELECTED" end if%>   value="0">Readjust</option>
								<option <% if ReworkType="1"  then response.write "SELECTED" end if%>    value="1">Tear down</option>
								<option  <% if ReworkType="2"  then response.write "SELECTED" end if%>  value="2">Slapping</option>
							</SELECT>
	</td>
    <td><span id="inner_SearchStartTime"></span></td>
    <td>
      <input name="txtfromdate" type="text" id="txtfromdate" value="<%=StartTime%>" size="10">
      <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.txtfromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
      </script>
		&nbsp;<span id="inner_SearchEndTime"></span>
	<input name="txttodate" type="text" id="txttodate" value="<%=EndTime%>" size="10">
	<script language=JavaScript type=text/javascript>
		function calendar2Callback(date, month, year)
		{
		document.all.txttodate.value=year + '-' + month + '-' + date
		}
		calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
	</script>
&nbsp; </td>
    </tr>
  <tr>
    <td height="20"><span id="inner_SearchOperatorCode"></span></td>
    <td height="20"><input name="txtOperatorCode" type="text" id="txtOperatorCode" value="<%=OperatorCode%>"></td>
    <td colspan="7"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right">
          <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('ReworkJob_Export.asp')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>&nbsp;</span></div></td>
      </tr>
    </table></td>
</tr>
 <tr>
      <td height="20" colspan="12"><!--#include virtual="/Components/PageSplit.asp" --></td>
    </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Part_Number"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_JobType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ReworkQty"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_GoodQty"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_RejectQty"></span></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StartTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_EndTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_OperatorCode"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Yield"></span></div></td>
 </tr>
<%
	i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<Tr>
  <td height="20"><div align="center"><%=i%></div></td>
  <td><div align="center">
	  <%if rs("STATUS")="1" then%>
	      <img src="\Images\Opened.gif" width="10" height="10">
	      <%else%>
	      <img src="\Images\Closed.gif" width="10" height="10">
	      <%end if%>
  </div></td>
  <td height="20"><div align="center"><%=rs("REWORK_JOB_NUMBER").value%></div></td>
  <td height="20"><div align="center">
  <%
  		set rs2=server.createobject("adodb.recordset")
		Job_Number=left(rs("REWORK_JOB_NUMBER").value,inStr(rs("REWORK_JOB_NUMBER").value,"-")-1)
		rs2.open "select PART_NUMBER_TAG from job where job_number='"&Job_Number&"'  and job_type='N' AND rownum=1",conn,1,3
		if(not rs2.eof) then
			response.write rs2(0).value
		else
			response.write "&nbsp;"
		end if
  %></div></td>
  
  <td><div align="center">
  	<% if rs("REWORK_TYPE").value="0" then response.write "Readjust"  end if 
	   if rs("REWORK_TYPE").value="1" then response.write "Teardown" end if  
	   if rs("REWORK_TYPE").value="2" then response.write "Slapping"  end if %></div>
  </td>
  <td ><div align="center"><%=rs("SUB_JOB_REWORK_QTY").value%></div></td>
  <td ><div align="center"><%=rs("REWORK_GOOD_QTY").value%><%if DBA=true AND  rs("STATUS")="2" then%><span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('ReworkQuantityUpdate.asp?jobnumber=<%=rs("REWORK_JOB_NUMBER")%>&REWORK_TYPE=<%=rs("REWORK_TYPE").value%>&SUBJOBNUMBER=<%=rs("SUB_JOB_NUMBER")%>','_new')">U</span><%end if%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("REWORK_REJECT_QTY").value%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("REWORK_START_TIME").value%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("REWORK_END_TIME").value%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("OPERATOR_CODE").value%></div></td>
  <td ><div align="center"><% if rs("STATUS")="2" then response.write (round(cdbl(rs("REWORK_GOOD_QTY").value)/cdbl(rs("SUB_JOB_REWORK_QTY").value),4)*100&"%") end if %>&nbsp;</div></td>
 </tr>
 
<%	
	i=i+1
	rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="12"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
<tr>
      <td height="20" colspan="12"><!--#include virtual="/Components/JobNote.asp" --></td>
    </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->