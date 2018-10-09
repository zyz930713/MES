<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetSingleQuantityAmount.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
family_name=request.QueryString("family_name")
factory_id=request.QueryString("factory_id")
line_name=request.QueryString("line_name")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Yield/DailyFinanceYieldYield/DailyFinanceYieldDetail.asp"
SQL="select NID,PREFIX from FINANCE_SERIES_GROUP where SERIES_GROUP_NAME='"&family_name&"' and FACTORY_ID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
id=rs("NID")
prefix=rs("PREFIX")
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy">Browse Recorded Daily Finance Yield of <%=family_name%> -- <%=line_name%> (<%=from_time%> - <%=to_time%>) </td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyFinanceYieldDetail_Export.asp?family_name=<%=family_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<tbody id="FG">
<%SQL="select PART_NUMBER,JOB_NUMBER,START_QUANTITY,START_VALUE,GOOD_QUANTITY,GOOD_VALUE,TWIN,TRANSACTION_DATE FROM BAR_REPORT.STOCK_TRANSACTION WHERE FINANCE_FAMILY_ID='"&id&"' and TRANSACTION_DATE>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and TRANSACTION_DATE<to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') order by JOB_NUMBER"
session("SQL")=SQL
rs.open SQL,conn,1,3%>
<tr>
  <td height="20" colspan="6" class="t-t-Borrow">Output Detail </td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Part Number</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number</div></td>
  <td class="t-t-Borrow"><div align="center">Start Value</div></td>
  <td class="t-t-Borrow"><div align="center">Output Value</div></td>
  <td class="t-t-Borrow"><div align="center">Date</div></td>
  </tr>
<%
i=1
total_start_quantity=0
total_start_value=0
total_good_quantity=0
total_good_value=0
if not rs.eof then
while not rs.eof
%>
<tr <%if rs("TWIN")="1" then%>class="t-b-Yellow"<%end if%>>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><%= rs("PART_NUMBER")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("JOB_NUMBER")%></div></td>
	 <td><div align="center"><%= rs("START_VALUE")%></div></td>
	 <td><div align="center"><%= rs("GOOD_VALUE")%></div></td>
	 <td><div align="center"><%= rs("TRANSACTION_DATE")%></div></td>
    </tr>
<%
if rs("TWIN")="1" then
twin_quantity=csng(rs("GOOD_QUANTITY"))
twin_amount=get_single_quantity_amount(rs("PART_NUMBER"),twin_quantity)
else
twin_quantity=0
twin_amount=0
end if
total_start_value=total_start_value+ccur(rs("START_VALUE"))-twin_amount
total_good_value=total_good_value+ccur(rs("GOOD_VALUE"))-twin_amount
i=i+1
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td><div align="center"><%= total_start_value%></div></td>
  <td><div align="center"><%= total_good_value%></div></td>
  <td>&nbsp;</td>
  </tr>
<%
end if
rs.close%>
</tbody>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tbody id="FG">
<%SQL="select PART_NUMBER,JOB_NUMBER,SERIES,QUANTITY,AMOUNT,SCRAP_DATE FROM BAR_REPORT.SCRAP_TRANSACTION WHERE SCRAP_DATE>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and SCRAP_DATE<to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') and SERIES in (select SERIES_NAME from FINANCE_SERIES where SERIES_GROUP_ID='"&id&"')"
'response.Write(SQL)
'response.End()
session("SQL1")=SQL
rs.open SQL,conn,1,3%>
<tr>
  <td height="20" colspan="6" class="t-t-Borrow">Scrap Detail </td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Part Number</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number</div></td>
  <td class="t-t-Borrow"><div align="center">Series</div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Value</div></td>
  <td class="t-t-Borrow"><div align="center">Date</div></td>
</tr>
<%
i=1
total_bad_quantity=0
total_bad_value=0
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><%= rs("PART_NUMBER")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("JOB_NUMBER")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("SERIES")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("AMOUNT")%></div></td>
     <td><div align="center"><%= rs("SCRAP_DATE")%></div></td>
</tr>
<%
total_bad_quantity=total_bad_quantity+ccur(rs("QUANTITY"))
total_bad_value=total_bad_value+ccur(rs("AMOUNT"))
i=i+1
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td><div align="center"><%= total_bad_value%></div></td>
  <td>&nbsp;</td>
</tr>
<%
end if
rs.close%>
</tbody>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->