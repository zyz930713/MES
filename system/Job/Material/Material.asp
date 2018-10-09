<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
server.ScriptTimeout=999999
response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
action_value=trim(request("action_value"))
fromdate=request("fromdate")
todate=request("todate")

if fromdate="" and jobnumber="" then
	fromdate=dateadd("d",-7,date())
end if
if todate="" then
	todate=dateadd("d",1,date())
end if

where=" AND J.START_TIME BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd') AND TO_DATE('"&todate&"','yyyy-mm-dd')"
if request("queryAction") <> "Query" then
	where=where&" and 1=2 "
end if
if action_value<>"" then
	where=where&" AND JA.ACTION_VALUE like '"&action_value&"%'"
end if

pagepara="&action_value="&action_value&"&fromdate="&fromdate&"&todate="&todate
pagename="/Job/Material/Material.asp"
set rs1=server.CreateObject("adodb.recordset")
SQL="SELECT JA.*,J.PART_NUMBER_TAG,J.LINE_NAME,J.START_TIME,"
SQL=SQL+" (SELECT STATION_NAME FROM STATION WHERE NID= JA.STATION_ID) STATION_NAME, "
SQL=SQL+" (SELECT ACTION_NAME FROM ACTION WHERE NID= JA.ACTION_ID) ACTION_NAME "
SQL=SQL+" FROM JOB_ACTIONS JA, JOB J WHERE JA.JOB_NUMBER =J.JOB_NUMBER AND JA.SHEET_NUMBER=J.SHEET_NUMBER AND JA.JOB_TYPE =J.JOB_TYPE "&where

session("SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
		job_numbers=job_numbers&rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)&","
		rs.movenext
	wend
	if job_numbers<>"" then
		job_numbers=left(job_numbers,len(job_numbers)-1)
	end if
	rs.movefirst
end if	
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
<!--#include virtual="/Language/Job/Material/Lan_Material.asp" -->
<script language="javascript">
function formyieldcheck()
{
	with(document.formYield)
	{
		if(action.selectedIndex==0)
		{
			alert("Please select which action to be convert!")
			return false;
		}
	}
}
function queryDate(){
	if(!form1.action_value.value){
		alert("Action Value cannot be blank. 数值不能为空。");
		return false;
	}
	document.form1.submit()	
}

</script>
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post" action="/Job/Material/Material.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchActionValue"></span></td>
    <td height="20"><input name="action_value" type="text" id="action_value" value="<%=action_value%>"></td>
    <td><span id="inner_SearchStartTime"></span></td>
    <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="queryDate();">
	<input type="hidden" name="queryAction" value="Query">
	</td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<form name="formYield" method="post" action="/Reports/Yield/JobYield/JobYield.asp" onSubmit="return formyieldcheck()">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="16%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="84%"><div align="right"><%if job_numbers<>"" then%>
		<select name="action" id="action">
		<option value="">-- Select Action --</option>
		  <%= getAction(true,"OPTION","",""," order by ACTION_NAME","","")%>
		</select>
		<input name="convert" type="submit" class="t-b-Yellow">
		<%end if%>&nbsp;&nbsp;<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Material_Export.asp?fromdate=<%=fromdate%>&todate=<%=todate%>')"><img src="/Images/EXCEL.gif" width="22" height="22" align="absmiddle"></span></div></td>
      </tr>
    </table></td>
</tr>
<input name="action_value" type="hidden" value="<%=action_value%>">
<input name="job_numbers" type="hidden" value="<%=job_numbers%>">
</form>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Station"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ActionName"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Value"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_RelativeActionName"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_RelativeValue"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td height="20"><div align="center" class="<%if rs("JOB_TYPE")="N" then%>LinkBlue<%else%>LinkGreen<%end if%>"><strong><span style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>','_blank')"><%=rs("JOB_NUMBER")%>-<%=replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)%></span></strong></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("LINE_NAME")%></div></td>
		<td><div align="center"><%= rs("STATION_NAME")%></div></td>
		<td><div align="center"><%= rs("ACTION_NAME")%></div></td>
		<td><div align="center"><%= rs("ACTION_VALUE")%></div></td>
		<%
		relative_action_name=""
		relative_action_value=""
'		SQL1="select A.ACTION_NAME,RJA.ACTION_VALUE from JOB_ACTIONS RJA inner join ACTION A on RJA.ACTION_ID=A.NID where RJA.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and RJA.SHEET_NUMBER="&rs("SHEET_NUMBER")&" and RJA.JOB_TYPE='"&rs("JOB_TYPE")&"' and RJA.ACTION_ID=(select NID from ACTION where RELATED_ACTION_ID='"&rs("ACTION_ID")&"')"
'		rs1.open SQL1,conn,1,3
'		if not rs1.eof then
'		relative_action_name=rs1("ACTION_NAME")
'		relative_action_value=rs1("ACTION_VALUE")
'		end if
'		rs1.close
		%>
		<td><div align="center"><%=relative_action_name%>&nbsp;</div></td>
		<td><div align="center"><%=relative_action_value%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center"><span id="inner_NoRecords"></span> </div></td>
  </tr>
<%end if
rs.close
set rs1=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->