<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TestDefectCode/TestDefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from TEST_DEFECTCODE where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/DefectCode/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/TestDefectCode/EditTestDefectCode1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a  Test Defect Code </td>
</tr>
  <tr>
    <td width="123" height="20">Defect Name  <span class="red">*</span> </td>
    <td width="642" height="20"><input name="defectname" type="text" id="defectname" value="<%=rs("DEFECT_NAME")%>" size="50"></td>
  </tr>
  <tr>
    <td height="20">Value Type   <span class="red">*</span> </td>
    <td height="20"><select name="value_type" id="value_type">
      <option value="1">Text</option>
      <option value="2">Decimal</option>
    </select></td>
  </tr>
  <tr>
    <td height="20">Scale  <span class="red">*</span> </td>
    <td height="20"><input name="scale" type="text" id="scale" value="0"></td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
&nbsp;
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->