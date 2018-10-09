<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/FinanceDual/DualCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from DUAL_SETTINGS where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/FinanceDual/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/FinanceDual/EditDual1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Daul Settings </td>
</tr>
<tr>
  <td width="153" height="20">Dual Model Name <span class="red">*</span> </td>
    <td width="601" height="20" class="red">
      <div align="left">
        <input name="Dual" type="text" id="Dual" value="<%=rs("DUAL_NAME")%>" size="30">
      </div></td>
    </tr>
<tr>
  <td height="20">Single Model Name 1 <span class="red">*</span></td>
  <td height="20"><input name="Single1" type="text" id="Single1" value="<%=rs("SINGLE1")%>" size="30"></td>
</tr>
<tr>
  <td height="20">Single Model Name 2 <span class="red">*</span></td>
  <td height="20"><input name="Single2" type="text" id="Single2" value="<%=rs("SINGLE2")%>" size="30"></td>
</tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
