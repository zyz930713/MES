<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
linename=request.QueryString("linename")

pagename="/Job/Shift/ShiftHistory.asp"
pagepara="&linename="&linename
SQL="select L.* from LINE_SHIFT L where L.LINE_NAME='"&linename&"' order by L.SHIFT_TIME desc"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/Shift/Lan_ShiftHistory.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="5"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td width="241" height="20" class="t-t-Borrow"><div align="center"><span id="inner_LineName"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_ShiftType"></span></div></td>
  <td width="235" height="20" class="t-t-Borrow"><div align="center"><span id="inner_ShiftPerson"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_ShiftTime"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(session("strpagenum")-1)*recordsize+i%>
  </div></td>
    <td height="20"><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td><div align="center"><%= rs("SHIFT_TYPE")%></div></td>
    <td height="20"><div align="center"><%= rs("SHIFT_PERSON")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("SHIFT_TIME")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="5"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->