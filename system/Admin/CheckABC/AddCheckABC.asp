<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/CheckABC/ABCCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Section/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/CheckABC/AddCheckABC1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>


<tr>
  <td height="20"><span id="td_TestName"></span> <span class="red">*</span></td>
  <td height="20"><input name="paramSF" type="text" id="paramSF"  value="" size="10" ></td>
</tr>
<tr>
  <td height="20">Low A <span class="red">*</span></td>
  <td height="20"><input name="LLA" type="text" id="LLA"  value="" size="40"></td>
</tr>
<tr>
  <td height="20">Up A <span class="red">*</span></td>
  <td height="20"><input name="ULA" type="text" id="ULA"  value="" size="40"></td>
</tr>
<tr>
  <td height="20">Low B <span class="red">*</span></td>
  <td height="20"><input name="LLB" type="text" id="LLB"  value="" size="40"></td>
</tr>
<tr>
  <td height="20">Up B <span class="red">*</span></td>
  <td height="20"><input name="ULB" type="text" id="ULB"  value="" size="40"></td>
</tr>
<tr>
  <td height="20"><span id="td_testType"></span> <span class="red">*</span></td>
  <td height="20"><input name="Product_Name" type="text" id="Product_Name"  value="" size="40"></td>
</tr>




  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="SolID" type="hidden" id="SolID" value="<%=SolID%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input name="Reset" type="reset" id="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>