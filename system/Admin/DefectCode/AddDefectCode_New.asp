<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
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
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/DefectCode/FormCheck.js" type="text/javascript"></script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onload="language(<%=session("language")%>);">
<form action="/Admin/DefectCode/AddDefectCode_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck_New()">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
</tr>
<tr> 
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
	<% =session("User") %></td>
</tr>
<tr>
  <td><span id="td_DefectCode"></span>  <span class="red">*</span> </td>
  <td class="red">
      <div align="left">
        <input name="defectcode" type="text" id="defectcode">
      </div></td>
    </tr>
  <tr>
    <td><span id="td_Name"></span>  <span class="red">*</span> </td>
    <td><input name="defectname" type="text" id="defectname" size="50"></td>
  </tr>
  <tr>
    <td><span id="td_CHName"></span>  <span class="red">*</span></td>
    <td><input name="chinesename" type="text" id="chinesename" size="50"></td>
  </tr>
  <tr>
    <td><span id="td_Factory"></span> <span class="red">*</span></td>
    <td><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
		<%FactoryRight ""%>
        <%= getFactory("OPTION","",factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input name="materialscount" type="hidden" id="materialscount">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input type="reset" name="btnReset" value="Reset">
&nbsp;
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->