<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsStockAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
code=trim(request("code"))
storetype=trim(request("storetype"))
section=trim(request("section"))
checkstatus=trim(request("checkstatus"))
fromdate=request("fromdate")
OBI_Status=request("OBI_Status")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" then
fromdate=date()
end if
if todate="" then
todate=dateadd("d",1,date())
end if

if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and lower(J.PART_NUMBER_TAG) like '%"&lcase(partnumber)&"%'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME) like '%"&lcase(line)&"%'"
end if
if section<>"" then
where=where&" and L.SECTION_ID='"&section&"'"
end if
if code<>"" then
where=where&" and J.STORE_CODE like '%"&code&"%'"
end if
if storetype<>"" then
where=where&" and J.STORE_TYPE='"&storetype&"'"
end if
if checkstatus<>"" then
where=where&" and J.RETEST_CHECK_STATUS='"&checkstatus&"'"
end if
if fromdate<>"" then
where=where&" and to_date(J.STORE_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.STORE_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
if OBI_Status<>"" then
where=where&" and J.ERP_JOB_CLOSE_STATUS='"&OBI_Status&"'"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&code="&code&"&storetype="&storetype&"&checkstatus="&checkstatus&"&fromdate="&fromdate&"&todate="&todate&"&OBI_Status="&OBI_Status
pagename="/Job/Store/StoreRecords/OldPreStoreRecords.asp"
FactoryRight "J."
SQL="select J.*,JM.ERP_JOB_STATUS from JOB_MASTER_STORE_PRE J inner join JOB_MASTER JM on J.JOB_NUMBER=JM.JOB_NUMBER inner join FACTORY F on J.FACTORY_ID=F.NID left join LINE L on J.LINE_NAME=L.LINE_NAME where J.store_code not like '%-0000' and J.job_number not IN(select job_number From job_master_scrap where scrap_code like '%0000%') "&where&factorywhereoutsideand&order
'response.write sql
session("SQL")=SQL
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
<script language="javascript">
function retestcheck(ob,thisvalue)
{
	if (ob.checked)
	{
	location.href="Retest_Check.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
	else
	{
	location.href="Retest_Uncheck.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
}
</script>
<!--#include virtual="/Language/Reports/Store/StoreRecords/Lan_OldPreStoreRecords.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/Job/Store/StoreRecords/OldPreStoreRecords.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span></td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>">
      <select name="checkstatus" id="checkstatus">
        <option value="">--Select Status--</option>
        <option value="1">Checked</option>
        <option value="0">Unchecked</option>
      </select></td>
    <td><span id="inner_SearchLine"></span></td>
    <td><input name="line" type="text" id="line" value="<%=line%>"></td>
    <td><span id="inner_SearchSection"></span></td>
    <td>      <select name="section">
	<option value="">--Select Section--</option>
        <%=getSection("OPTION",section,"",NULL,NULL)%>
      </select></td>
    <td><span id="inner_SearchStoreTime"></span></td>
    <td><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
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
    </tr>
  <tr>
    <td height="20"><span id="inner_SearchCode"></span></td>
    <td height="20"><input name="code" type="text" id="code" value="<%=code%>"></td>
    <td><span id="inner_SearchStoreType"></span></td>
    <td><select name="storetype" id="storetype">
      <option value="">--Select Type--</option>
      <option value="N">Normal</option>
      <option value="R">Rework</option>
    </select>    </td>
    <td><span id="inner_SearchCheckStatus"></span></td>
    <td>&nbsp;</td>
    <td>OBI Status</td>
    <td><select name="OBI_Status" id="OBI_Status">
        <option value="">-- Status --</option>
        <option value="NEW" <%IF OBI_Status="NEW" THEN%>SELECTED<%END IF%>>NEW</option>
        <option value="PENDING" <%IF OBI_Status="PENDING" THEN%>SELECTED<%END IF%>>PENDING</option>
        <option value="SUCCESS" <%IF OBI_Status="SUCCESS" THEN%>SELECTED<%END IF%>>SUCCESS</option>
        <option value="FAILURE" <%IF OBI_Status="FAILURE" THEN%>SELECTED<%END IF%>>FAILURE</option>
        <option value="VALIDATION_ERROR" <%IF OBI_Status="VALIDATION_ERROR" THEN%>SELECTED<%END IF%>>VALIDATION_ERROR</option>
      </select>
    </td>
    <td>&nbsp;</td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><span id="inner_Browse"></span><a href="/Job/Store/StoreRecords/PreStoreRecords.asp" target="_blank" class="white">去新的系统</a></td>
</tr>
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('StoreRecords_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="18"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if retestchecker=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Check"></span></div></td>  <%end if%>

  <td class="t-t-Borrow"><div align="center">Txn  ID</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Code"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_InputQuantity"></span></div></td>
<!--  <td class="t-t-Borrow"><div align="center"><span id="inner_InspectQuantity"></span></div></td>
-->  <td class="t-t-Borrow"><div align="center"><span id="inner_StoreQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_OntimeYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_StoreTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StoreType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Note"></span></div></td>
  <td class="t-t-Borrow"><div align="center">OBI Request Time</div></td>
  <td class="t-t-Borrow"><div align="center">ERP Job Status</div></td>
  <td class="t-t-Borrow"><div align="center">OBI Status </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_SubJobs"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 

%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
	  	<%if retestchecker=true then%>
		<td><div align="center">
		  <input name="retestcheck<%=i%>" type="checkbox" id="retestcheck<%=i%>" value="<%=rs("NID")%>" <%if rs("RETEST_CHECK_STATUS")="1" then%>checked<%end if%> onClick="retestcheck(this,this.value)">
	    </div></td>
		<%end if%>
		<td><div align="center"><%= rs("TRANSACTION_ID")%>&nbsp;</div></td>
		<td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
		<td><div align="center"><%if checker=false then%><% =rs("LINE_NAME")%><%else%>
		  <input name="line_name<%=i%>" type="text" value="<%=rs("LINE_NAME")%>" size="4">
		  <span class="red" style="cursor:hand" onClick="javascript:location.href='LineUpdate.asp?nid=<%=rs("NID")%>&new_line='+document.all.line_name<%=i%>.value+'&path=<%=path%>&query=<%=query%>'">U</span>
		  <%end if%>
		</div></td>
		<td><div align="center"><%=rs("STORE_CODE")%></div></td>
		<td><div align="center"><%if checker=false then%><% =rs("INPUT_QUANTITY")%><%else%>
		  <input name="input_quanity<%=i%>" type="text" value="<%=rs("INPUT_QUANTITY")%>" size="3">
		  <span class="red" style="cursor:hand" onClick="javascript:location.href='InputUpdate.asp?nid=<%=rs("NID")%>&new_input='+document.all.input_quanity<%=i%>.value+'&path=<%=path%>&query=<%=query%>'">U</span>
		  <%end if%></div></td>
<!--		<td><div align="center"><%'=rs("INSPECT_QUANTITY")%>&nbsp;</div></td>
-->		<td><div align="center"><%=rs("STORE_QUANTITY")%></div></td>
		<td><div align="center"><%=formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
		<td><div align="center"><%if checker=false then%><% =formatdate(rs("STORE_TIME"),application("longdateformat"))%><%else%>
		  <input name="store_time<%=i%>" type="text" id="store_time<%=i%>" value="<% =formatdate(rs("STORE_TIME"),application("longdateformat"))%>" size="19">
		  <span class="red" style="cursor:hand" onClick="javascript:location.href='TimeUpdate.asp?nid=<%=rs("NID")%>&new_time='+document.all.store_time<%=i%>.value+'&path=<%=path%>&query=<%=query%>'">U</span>
		  <%end if%>
		</div></td>
		<td><div align="center"><%=rs("STORE_TYPE")%></div></td>
		<td><div align="center"><%=rs("NOTE")%>&nbsp;</div></td>
		<td><div align="center"><%=rs("ERP_LAST_UPDATE_TIME")%>&nbsp;</div></td>
		<td><div align="center"><%=rs("ERP_JOB_STATUS")%>&nbsp;</div></td>
		<td><div align="center"><%=rs("ERP_JOB_CLOSE_STATUS")%>&nbsp;</div></td>
		<td><div align="left"><%=formatlongstringbreak(rs("SUB_JOB_NUMBERS"),"<br>",30)%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="18"><div align="center"><span id="inner_NoRecords">No Records</span></div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->