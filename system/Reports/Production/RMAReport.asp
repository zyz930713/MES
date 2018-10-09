<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->

<%
set rssum=server.createobject("adodb.recordset")
fromdate=request("fromdate")
todate=request("todate")
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
todate=request("todate")
if isnull(todate) or todate=""  then	
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if

RMA_ID=request("RMA_ID")
BOX_ID=request("BOX_ID")
modelstatus=request("status")
factory=request("factory")

strWhere=" WHERE RMA_TIME BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd hh24:mi:ss') and TO_DATE('"&todate&" 23:59:59','yyyy-mm-dd hh24:mi:ss') "

if RMA_ID<>"" then
strWhere=strWhere&" and RMA_ID like '%"&RMA_ID&"%'"
end if

pagepara="&RMA_ID="&RMA_ID&"&fromdate="&fromdate&"&todate="&todate

SQL="select * from RMA_LIST "&strWhere&" order by RMA_TIME"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">

</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Reports/Production/RMAReport.asp" method="get" name="form1" target="_self">
               


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_RMAID"></span></td>
    <td height="20"><input name="RMA_ID" type="text" id="RMA_ID" value="<%=RMA_ID%>"></td>
    <td><span id="inner_WHRECTime"></span></td>
    <td>&nbsp;<span id="inner_SearchFrom"></span>&nbsp;<input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" readonly size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>"  readonly="true" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">	</td>
   </tr>
     
</table>




</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Packing_Plan/addPacking_Plan.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td> 
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_RMAID"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_BoxId"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Quantity"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_WHRECUser"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_WHRECTime"></span></div></td>
 
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
 
    <td height="20"><div align="center"><%= rs("RMA_ID") %></div></td>
	<td><%= rs("BOX_ID") %></td>

	<td><%= rs("PACKED_QTY")%>&nbsp;</td>
    <td><%= rs("RMA_USER")%>&nbsp;</td>
    <td><%= rs("RMA_TIME")%>&nbsp;</td>
	
   
  </tr>
<%
i=i+1
rs.movenext
wend




sqlsum="select sum(qtysum) as qtysum from (select  sum(packed_qty) as qtysum from RMA_LIST "&strWhere&" group by RMA_id)"

'response.Write(sqlsum)
rssum.open sqlsum,conn,1,3
if  not rssum.bof then
qtysum=rssum("qtysum")
end if
rssum.close%>
<tr>
    <td height="20" colspan="14" align="center"><span id="inner_Quantity"></span>:<%=qtysum%></td>
  </tr>
<%


else
%>
  <tr>
    <td height="20" colspan="13" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->