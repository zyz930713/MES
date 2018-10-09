<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSubPart.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from DEFECTCODE_New where NID='"&id&"' order by defect_code"
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
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onload="language(<%=session("language")%>);">
<form action="/Admin/DefectCode/EditDefectCode_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
</tr>
<tr> 
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
	<% =session("User") %></td>
</tr>
<tr>
  <td width="121" height="20"><span id="td_DefectCode"></span>  <span class="red">*</span> </td>
  <td width="597" height="20" class="red">
      <div align="left">
        <input name="defectcode" type="text" id="defectcode" value="<%=rs("DEFECT_CODE")%>" size="50">
      </div></td>
    </tr>
  <tr>
    <td height="20"><span id="td_Name"></span> <span class="red">*</span> </td>
    <td height="20"><input name="defectname" type="text" id="defectname" value="<%=rs("DEFECT_NAME")%>" size="50"></td>
  </tr>
  <tr>
    <td height="20"><span id="td_CHName"></span>  <span class="red">*</span></td>
    <td height="20"><input name="chinesename" type="text" id="chinesename" value="<%=rs("DEFECT_CHINESE_NAME")%>" size="50"></td>
  </tr>
  <tr>
    <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
		<%FactoryRight ""%>
        <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
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
<!--#include virtual="/Functions/TableControl.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->