<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsOBI.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
planer=trim(request("planer"))
progress=request("progress")
fromdate=request("fromdate")
todate=request("todate")
close_fromdate=request("close_fromdate")
close_todate=request("close_todate")
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
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&planer="&planer&"&progress="&progress&"&fromdate="&fromdate&"&todate="&todate&"&close_fromdate="&close_fromdate&"&close_todate="&close_todate
pagename="/Job/Job/Job.asp"
FactoryRight "J."
SQL="SELECT 1 AS TXN_SOURCE_TBL,J.NID,J.TRANSACTION_ID,J.JOB_NUMBER,J.PART_NUMBER_TAG,J.LINE_NAME,J.ALLOW_ERP_RESUBMIT,J.STORE_QUANTITY AS COMPLETE_QTY,0 AS SCRAP_QTY,0 AS STORE_RETURN_QTY,NULL AS SCRAP_ACCOUNT,NULL SCRAP_REASON_ID,J.STORE_TIME AS TXN_TIME,J.ERP_LAST_UPDATE_TIME AS ERP_COMPLETE_TIME,J.ERP_JOB_CLOSE_STATUS AS ERP_CLOSE_STATUS,J.ERP_ERROR_CODE,J.ERP_ERROR_EXPLANATION,JM.ERP_JOB_STATUS FROM JOB_MASTER_STORE J INNER JOIN JOB_MASTER JM ON J.JOB_NUMBER=JM.JOB_NUMBER WHERE J.TRANSACTION_ID IS NOT NULL AND J.MESSAGE_STATUS IN ('SYSTEM_ERROR', 'UNKNOWN_ERROR') AND J.ERP_JOB_CLOSE_STATUS<>'SUCCESS' AND NVL(JM.OBI_ENABLED, 'Y') = 'Y' "
SQL=SQL&"UNION ALL SELECT 2 AS TXN_SOURCE_TBL,J.NID,J.TRANSACTION_ID,TO_NCHAR(J.JOB_NUMBER),TO_NCHAR(J.PART_NUMBER_TAG),TO_NCHAR(J.LINE_NAME),J.ALLOW_ERP_RESUBMIT,0 AS COMPLETE_QTY,J.SCRAP_QUANTITY AS SCRAP_QTY,0 AS STORE_RETURN_QTY,J.SCRAP_ACCOUNT,J.SCRAP_REASON_ID,J.SCRAP_TIME AS TXN_TIME,J.ERP_LAST_UPDATE_TIME AS ERP_COMPLETE_TIME,J.ERP_JOB_CLOSE_STATUS AS ERP_CLOSE_STATUS,J.ERP_ERROR_CODE,J.ERP_ERROR_EXPLANATION,JM.ERP_JOB_STATUS FROM JOB_MASTER_SCRAP J INNER JOIN JOB_MASTER JM ON J.JOB_NUMBER=JM.JOB_NUMBER WHERE J.TRANSACTION_ID IS NOT NULL AND J.MESSAGE_STATUS IN ('SYSTEM_ERROR', 'UNKNOWN_ERROR') AND J.ERP_JOB_CLOSE_STATUS <> 'SUCCESS' AND NVL(JM.OBI_ENABLED, 'Y') = 'Y' "
SQL=SQL&"UNION ALL SELECT 3 AS TXN_SOURCE_TBL,J.NID,J.TRANSACTION_ID,J.JOB_NUMBER,J.PART_NUMBER_TAG,J.LINE_NAME,J.ALLOW_ERP_RESUBMIT,0 AS COMPLETE_QTY,0 AS SCRAP_QTY,J.OLD_STORE_QUANTITY - J.NEW_STORE_QUANTITY AS STORE_RETURN_QTY,NULL as SCRAP_ACCOUNT,NULL as SCRAP_REASON_ID,J.CHANGE_TIME AS TXN_TIME,J.ERP_LAST_UPDATE_TIME AS ERP_COMPLETE_TIME,J.ERP_JOB_CLOSE_STATUS AS ERP_CLOSE_STATUS,J.ERP_ERROR_CODE,J.ERP_ERROR_EXPLANATION,JM.ERP_JOB_STATUS FROM JOB_MASTER_STORE_CHANGE J INNER JOIN JOB_MASTER JM ON J.JOB_NUMBER=JM.JOB_NUMBER WHERE J.TRANSACTION_ID IS NOT NULL AND J.MESSAGE_STATUS IN ('SYSTEM_ERROR', 'UNKNOWN_ERROR') AND J.ERP_JOB_CLOSE_STATUS <> 'SUCCESS' AND NVL(JM.OBI_ENABLED, 'Y') = 'Y' "
SQL=SQL&"UNION ALL SELECT 4 AS TXN_SOURCE_TBL,J.NID,J.TRANSACTION_ID,TO_NCHAR(J.JOB_NUMBER),TO_NCHAR(J.PART_NUMBER_TAG),TO_NCHAR(J.LINE_NAME),J.ALLOW_ERP_RESUBMIT,0 AS COMPLETE_QTY,0 AS SCRAP_QTY,J.OLD_SCRAP_QUANTITY - J.NEW_SCRAP_QUANTITY AS STORE_RETURN_QTY,JMS.SCRAP_ACCOUNT,JMS.SCRAP_REASON_ID,J.CHANGE_TIME AS TXN_TIME,J.ERP_LAST_UPDATE_TIME AS ERP_COMPLETE_TIME,J.ERP_JOB_CLOSE_STATUS AS ERP_CLOSE_STATUS,J.ERP_ERROR_CODE,J.ERP_ERROR_EXPLANATION,JM.ERP_JOB_STATUS FROM JOB_MASTER_SCRAP_CHANGE J INNER JOIN JOB_MASTER JM ON J.JOB_NUMBER=JM.JOB_NUMBER INNER JOIN JOB_MASTER_SCRAP JMS ON J.SCRAP_NID=JMS.NID WHERE J.TRANSACTION_ID IS NOT NULL AND J.MESSAGE_STATUS IN ('SYSTEM_ERROR', 'UNKNOWN_ERROR') AND J.ERP_JOB_CLOSE_STATUS <> 'SUCCESS' AND NVL(JM.OBI_ENABLED, 'Y') = 'Y'"&where&factorywhereoutsideand
'response.Write(SQL)
'response.End()
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
<!--#include virtual="/Language/OBI/Lan_JobClosing.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/OBI/JobSysClosing.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td rowspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
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
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><!--<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Job_Export.asp?fromdate=<%'=fromdate%>&todate=<%'=todate%>')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>-->&nbsp;</div></td>
      </tr>
    </table>	</td>
</tr>
<tr>
  <td height="20" colspan="12"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Txn  ID</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StartQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Return Quantity </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.FIRST_STORE_TIME	&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_ClosingTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.FIRST_STORE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">ERP Job Status</div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Error"></span> </div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td><div align="center"><%=rs("TRANSACTION_ID")%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/Job.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%=rs("JOB_NUMBER")%></a></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("LINE_NAME")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("COMPLETE_QTY")%></div></td>
	    <td><div align="center"><%= rs("SCRAP_QTY")%></div></td>
      <td><div align="center"><%= rs("STORE_RETURN_QTY")%></a></div></td>
		<td><div align="center">
		  <% =formatdate(rs("ERP_COMPLETE_TIME"),application("longdateformat"))%>
	    &nbsp;</div></td>
		<td><div align="center">
            <% =rs("ERP_JOB_STATUS")%>&nbsp;
        </div></td>
		<td><div align="center"><% =rs("ERP_ERROR_EXPLANATION")%>&nbsp;</div></td>
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
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->