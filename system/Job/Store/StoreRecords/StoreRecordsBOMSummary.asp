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
fromdate=request("fromdate")
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

if fromdate<>"" then
where=where&" and to_date(J.STORE_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.STORE_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&fromdate="&fromdate&"&todate="&todate
pagename="/Job/Store/StoreRecords/StoreRecordsBOMSummary.asp"
SQL="select TO_CHAR(J.STORE_TIME,'yyyy-MM-dd') as STORE_TIME,SUM(INPUT_QUANTITY) as INPUT_QUANTITY,SUM(STORE_QUANTITY) as STORE_QUANTITY,JM.PART_NUMBER_TAG,JM.BOM_LABEL from JOB_MASTER_STORE J inner join JOB_MASTER JM on J.JOB_NUMBER=JM.JOB_NUMBER where J.JOB_NUMBER is not null "&where&"  GROUP BY TO_CHAR(J.STORE_TIME,'yyyy-MM-dd'),JM.PART_NUMBER_TAG,JM.BOM_LABEL"
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

<!--#include virtual="/Language/Reports/Store/StoreRecords/Lan_StoreRecordsBOMSummary.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/Job/Store/StoreRecords/StoreRecordsBOMSummary.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchStoreTime"></span></td>
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
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; <input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('StoreRecords_Export2.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Date</div></td>

  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">BOM</div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_InputQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StoreQuantity"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td><div align="center"><a href="StoreRecords.asp?fromdate=<% =rs("STORE_TIME")%>&todate=<% =dateadd("d",1,rs("STORE_TIME"))%>&partnumber=<% =rs("PART_NUMBER_TAG")%>&BOM=<%=rs("BOM_LABEL")%>"><% =rs("STORE_TIME")%></a></div></td>
		<td><div align="center"><% =rs("PART_NUMBER_TAG")%></div></td>
		<td><div align="center"><%=rs("BOM_LABEL")%>&nbsp;</div></td>
		<td><div align="center">
		  <% =rs("INPUT_QUANTITY")%>
		</div></td>
		<td><div align="center"><%=rs("STORE_QUANTITY")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="6"><div align="center"><span id="inner_NoRecords">No Records</span></div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->