<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsOBI.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
planer=trim(request("planer"))
progress=request("progress")
fromdate=request("fromdate")
todate=request("todate")
close_fromdate=request("close_fromdate")
close_todate=request("close_todate")
resubmittable=request("resubmittable")
errorcode=request("errorcode")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" and jobnumber="" and close_fromdate="" and close_todate="" then
'fromdate="2007-1-1"
fromdate=dateadd("d",-7,date())
end if
if todate="" then
'todate=dateadd("d",1,date())
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
if planer<>"" then
where=where&" and lower(J.CREATE_NAME) like '%"&lcase(planer)&"%'"
end if
if progress="0" then
where=where&" and J.STORE_STATUS=0"
elseif progress="1" then
where=where&" and J.STORE_STATUS=1"
end if
if close_fromdate<>"" then
where=where&" and J.ERP_LAST_UPDATE_TIME>=to_date('"&close_fromdate&"','yyyy-mm-dd')"
end if
if close_todate<>"" then
where=where&" and J.ERP_LAST_UPDATE_TIME<=to_date('"&close_todate&"','yyyy-mm-dd')"
end if
if resubmittable<>"" then
where=where&" and J.ALLOW_ERP_RESUBMIT='"&resubmittable&"'"
end if
if errorcode<>"" then
where=where&" and J.ERP_ERROR_CODE='"&errorcode&"'"
end if

pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&planer="&planer&"&progress="&progress&"&fromdate="&fromdate&"&todate="&todate&"&close_fromdate="&close_fromdate&"&close_todate="&close_todate&"&resubmittable="&resubmittable&"&errorcode="&errorcode
FactoryRight "J."
SQL="select J.NID,J.TRANSACTION_ID,J.JOB_NUMBER,J.PART_NUMBER_TAG,J.LINE_NAME,J.OLD_SCRAP_QUANTITY,J.NEW_SCRAP_QUANTITY,J.CHANGE_CODE ,J.CHANGE_TIME ,NVL(J.ALLOW_ERP_RESUBMIT,'Y') AS ALLOW_ERP_RESUBMIT,J.ERP_LAST_UPDATE_TIME,J.ERP_ERROR_CODE,J.ERP_ERROR_EXPLANATION,JM.ERP_JOB_STATUS from JOB_MASTER_SCRAP_CHANGE J inner join JOB_MASTER JM on J.JOB_NUMBER=JM.JOB_NUMBER left join FACTORY F on J.FACTORY_ID=F.NID where TRANSACTION_ID IS NOT NULL AND ERP_JOB_CLOSE_STATUS IN ('VALIDATION_ERROR', 'FAILURE') AND NVL(JM.OBI_ENABLED, 'Y') = 'Y'"&where&factorywhereoutsideand&order
'response.Write(SQL)
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
<!--#include virtual="/Language/OBI/Lan_JobScrapReturn.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/OBI/JobScrapReturn.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Resubmittable</td>
    <td><select name="resubmittable" id="resubmittable">
      <option value="">--Select--</option>
      <option value="Y" <%if resubmittable="Y" then%>selected<%end if%>>Y</option>
      <option value="N" <%if resubmittable="N" then%>selected<%end if%>>N</option>
    </select></td>
    </tr>
  <tr>
    <td height="20"><span id="inner_SearchLine"></span></td>
    <td height="20"><input name="line" type="text" id="line" value="<%=line%>" size="6"></td>
    <td><span id="inner_SearchJobCloseTime"></span></td>
    <td><input name="close_fromdate" type="text" id="close_fromdate" value="<%=close_fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.close_fromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="close_todate" type="text" id="close_todate" value="<%=close_todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar4Callback(date, month, year)
	{
	document.all.close_todate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
    </script></td>
    <td>Error Code</td>
    <td><select name="errorcode" id="errorcode">
      <option value="">--Select--</option>
      <option value="ORACLE_API_ERROR" <%if errorcode="ORACLE_API_ERROR" then%>selected<%end if%>>ORACLE_API_ERROR</option>
      <option value="XX_ASSEMBLY_LOT_INFO_NOT_NULL" <%if errorcode="XX_ASSEMBLY_LOT_INFO_NOT_NULL" then%>selected<%end if%>>XX_ASSEMBLY_LOT_INFO_NOT_NULL</option>
      <option value="XX_COMPONENT_LOT_INFO_NOT_NULL" <%if errorcode="XX_COMPONENT_LOT_INFO_NOT_NULL" then%>selected<%end if%>>XX_COMPONENT_LOT_INFO_NOT_NULL</option>
      <option value="XX_INVALID_ASSEMBLY_LOT_INFO" <%if errorcode="XX_INVALID_ASSEMBLY_LOT_INFO" then%>selected<%end if%>>XX_INVALID_ASSEMBLY_LOT_INFO</option>
      <option value="XX_INVALID_ASSEMBLY_LOT_QTY" <%if errorcode="XX_INVALID_ASSEMBLY_LOT_QTY" then%>selected<%end if%>>XX_INVALID_ASSEMBLY_LOT_QTY</option>
      <option value="XX_INVALID_COMPLETE_QTY" <%if errorcode="XX_INVALID_COMPLETE_QTY" then%>selected<%end if%>>XX_INVALID_COMPLETE_QTY</option>
      <option value="XX_INVALID_COMPONENT_LOT_NUMBER" <%if errorcode="XX_INVALID_COMPONENT_LOT_NUMBER" then%>selected<%end if%>>XX_INVALID_COMPONENT_LOT_NUMBER</option>
      <option value="XX_INVALID_COMPLETION_SUBINV" <%if errorcode="XX_INVALID_COMPLETION_SUBINV" then%>selected<%end if%>>XX_INVALID_COMPLETION_SUBINV</option>
      <option value="XX_INVALID_INK_DOT_QTY" <%if errorcode="XX_INVALID_INK_DOT_QTY" then%>selected<%end if%>>XX_INVALID_INK_DOT_QTY</option>
      <option value="XX_INVALID_JOB_NUMBER" <%if errorcode="XX_INVALID_JOB_NUMBER" then%>selected<%end if%>>XX_INVALID_JOB_NUMBER</option>
      <option value="XX_INVALID_JOB_STATUS" <%if errorcode="XX_INVALID_JOB_STATUS" then%>selected<%end if%>>XX_INVALID_JOB_STATUS</option>
      <option value="XX_INVALID_ORG_ID" <%if errorcode="XX_INVALID_ORG_ID" then%>selected<%end if%>>XX_INVALID_ORG_ID</option>
      <option value="XX_INVALID_SCRAP_ACCOUNT" <%if errorcode="XX_INVALID_SCRAP_ACCOUNT" then%>selected<%end if%>>XX_INVALID_SCRAP_ACCOUNT</option>
      <option value="XX_INVALID_SCRAP_QTY" <%if errorcode="XX_INVALID_SCRAP_QTY" then%>selected<%end if%>>XX_INVALID_SCRAP_QTY</option>
      <option value="XX_INVALID_SCRAP_REASON" <%if errorcode="XX_INVALID_SCRAP_REASON" then%>selected<%end if%>>XX_INVALID_SCRAP_REASON</option>
      <option value="XX_NEGATIVE_ONHAND_QTY" <%if errorcode="XX_NEGATIVE_ONHAND_QTY" then%>selected<%end if%>>XX_NEGATIVE_ONHAND_QTY</option>
      <option value="XX_NOT_ENOUGH_QTY" <%if errorcode="XX_NOT_ENOUGH_QTY" then%>selected<%end if%>>XX_NOT_ENOUGH_QTY</option>
      <option value="XX_INVALID_SCRAP_QTY" <%if errorcode="XX_INVALID_SCRAP_QTY" then%>selected<%end if%>>XX_INVALID_SCRAP_QTY</option>
      <option value="XX_OBI_NOT_ALLOWED" <%if errorcode="XX_OBI_NOT_ALLOWED" then%>selected<%end if%>>XX_OBI_NOT_ALLOWED</option>
      <option value="XX_ORG_PERIOD_NOT_OPEN" <%if errorcode="XX_ORG_PERIOD_NOT_OPEN" then%>selected<%end if%>>XX_ORG_PERIOD_NOT_OPEN</option>
      <option value="XX_START_QTY_MISMATCH" <%if errorcode="XX_START_QTY_MISMATCH" then%>selected<%end if%>>XX_START_QTY_MISMATCH</option>
      <option value="XX_UNKNOWN_VAL_ERROR" <%if errorcode="XX_UNKNOWN_VAL_ERROR" then%>selected<%end if%>>XX_UNKNOWN_VAL_ERROR</option>
    </select>      <img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><!--<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Job_Export.asp?fromdate=<%'=fromdate%>&todate=<%'=todate%>')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>-->&nbsp;</div></td>
      </tr>
    </table>	</td>
</tr>
<tr>
  <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Txn  ID</div></td>
  <%if DBA=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_OBIResubmit"></span></div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Select"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_OldScrapQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_NewScrapQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Change By </div></td>
  <td class="t-t-Borrow"><div align="center">Change Time</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.ERP_LAST_UPDATE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_ClosingTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.ERP_LAST_UPDATE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">ERP Job Status</div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Error"></span>Error Code</div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Error"></span></div></td>
  </tr>
<form action="/OBI/JobScrapReturn1.asp" method="post" name="checkform" id="checkform">
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
	  <td><div align="center"><%=rs("TRANSACTION_ID")%></div></td>
<%if OBI=true then%>
		<td><div align="center"><%if rs("ALLOW_ERP_RESUBMIT")="Y" then%><img src="/Images/Enabled.gif" width="10" height="10" style="cursor:hand" onClick="javascript:location.href='DisableJobScrapReturnResubmit.asp?p_nid=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>'"><%else%><img src="/Images/Disabled.gif" width="10" height="10" style="cursor:hand" onClick="javascript:location.href='EnableJobScrapReturnResubmit.asp?p_nid=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>'"><%end if%>&nbsp;</div></td>		<%end if%>
		<td>
		  <div align="center">
		    <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1" <%if rs("ALLOW_ERP_RESUBMIT")<>"Y" then%>disabled<%end if%>>
	        <input name="nid<%=i%>" type="hidden" id="nid<%=i%>" value="<%=rs("NID")%>">
		  </div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/Job.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%=rs("JOB_NUMBER")%></a></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("LINE_NAME")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("OLD_SCRAP_QUANTITY")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("NEW_SCRAP_QUANTITY")%>&nbsp;</div></td>
	    <td><div align="center"><%= rs("CHANGE_CODE")%></div></td>
	    <td><div align="center"><%= rs("CHANGE_TIME")%></div></td>
      <td><div align="center">
		  <% =formatdate(rs("ERP_LAST_UPDATE_TIME"),application("longdateformat"))%>
	    &nbsp;</div></td>
		<td><div align="center">
            <% =rs("ERP_JOB_STATUS")%>
        </div></td>
		<td><div align="center">
		  <% =rs("ERP_ERROR_CODE")%>
		  &nbsp;</div></td>
		<td><div align="center"><% =rs("ERP_ERROR_EXPLANATION")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend%>
 <tr>
    <td height="20" colspan="16"><div align="center">
	<input name="path" type="hidden" id="path" value="<%=path%>">
	<input name="query" type="hidden" id="query" value="<%=query%>">
	<input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
     <input name="CheckAll" type="button" id="CheckAll" value="Check All" onClick="checkall()">
&nbsp;
<input name="UncheckAll" type="button" id="UncheckAll" value="Uncheck All" onClick="uncheckall()">&nbsp;
<input name="Resubmit" type="submit" id="Resubmit" value="重新提交">&nbsp;<input name="Reset" type="button" id="Reset" value="重置"></div></td>
  </tr></form> 
<%
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->