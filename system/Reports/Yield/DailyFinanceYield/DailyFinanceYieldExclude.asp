<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory_id=request.QueryString("factory_id")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Yield/DailyFinanceYieldYield/DailyFinanceYieldExclude.asp"
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
  <td height="20" colspan="13" class="t-c-greenCopy">Browse Recorded Daily Finance Yield of <%=family_name%> -- <%=line_name%> (<%=from_time%> - <%=to_time%>) </td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyFinanceYieldDetail_Export.asp?family_name=<%=family_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<tbody id="FG">
<%SQL="select PART_NUMBER,REFERENCE,QUANTITY,AMOUNT FROM BAR_REPORT.SCRAP_TRANSACTION WHERE SCRAP_DATE BETWEEN to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') and FACTORY='"&factory_id&"' and FINANCE_FAMILY_ID is null"
'response.Write(SQL)
'response.End()
session("SQL")=SQL
rs.open SQL,conn,1,3%>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Reference</div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Value </div></td>
  </tr>
<%
i=1
total_ma_bad_quantity=0
total_ma_bad_value=0
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><%= rs("PART_NUMBER")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("REFERENCE")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("QUANTITY")%></div></td>
	 <td><div align="center"><%= rs("AMOUNT")%></div></td>
    </tr>
<%
total_ma_bad_quantity=total_ma_bad_quantity+ccur(rs("QUANTITY"))
total_ma_bad_value=total_ma_bad_value+ccur(rs("AMOUNT"))
i=i+1
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td><div align="center"><%= total_ma_bad_quantity%></div></td>
  <td><div align="center"><%= total_ma_bad_value%></div></td>
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