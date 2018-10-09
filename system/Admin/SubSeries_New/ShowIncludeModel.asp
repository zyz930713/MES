<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
NID=request.querystring("id")
SQL="SELECT ITEM_NAME,DESCRIPTION FROM PRODUCT_MODEL WHERE SERIES_ID='"& NID &"' ORDER BY ITEM_NAME "

rs.open SQL,conn,1,3

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
  	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Description"></span></div></td>
  </tr>
 <%
while not rs.eof 
%>
    <tr>
     <td height="20"><div align="center"><%= rs("ITEM_NAME") %></div></td>
	 <td height="20"><%= rs("DESCRIPTION") %></td>
    </tr>
    <%
rs.movenext
wend
rs.close
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->