<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
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
<script language="JavaScript" src="/Admin/GlueWIRE/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/GlueWire/AddGlueWire1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td height="20"><span id="inner_PartNumber"></span> <span class="red">*</span></td>
  <td height="20"> <input name="ITEM_NAME" type="text" id="ITEM_NAME"  value="" size="15"></td>
</tr>
<tr>
  <td height="20"><span id="inner_SupplierName"></span> <span class="red">*</span></td>
  <td height="20"><input name="SUPPLIER_NAME" type="text" id="SUPPLIER_NAME"  value="" size="10"></td>
</tr>
<tr>
  <td height="20"><span id="inner_StationName"></span> <span class="red">*</span></td>
  <td height="20"><input name="STATION_DESCRIPTION" type="text" id="STATION_DESCRIPTION"  value="" size="40"></td>
</tr>
<tr>
  <td height="20"><span id="inner_SubSeries"></span> <span class="red">*</span></td>
  <td height="20"><input name="STATION_DESCRIPTION_EN" type="text" id="STATION_DESCRIPTION_EN"  value="" size="40" readonly> 
<select name="select11" onChange="(document.form1.STATION_DESCRIPTION_EN.value=this.options[this.selectedIndex].value)">
<option >��ѡ��</option>
<%
set rs_s=server.createobject("adodb.recordset")
rs_s.open "select PRODUCT from RPT_DAILY_TARGET  group by PRODUCT ",conn,1,3
%>
<%
while not rs_s.eof%>
<option value="<%=rs_s("PRODUCT")%>"><%=rs_s("PRODUCT")%></option>
<%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
</select>
      </td>
</tr>



  <tr>
    <td height="20" colspan="2"><div align="center">
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->