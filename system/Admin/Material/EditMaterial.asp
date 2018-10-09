<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Material/MaterialCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from MATERIAL where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Material/FormCheck.js" type="text/javascript"></script>
</head>

<body>

<form action="/Admin/Material/EditMaterial1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a  Material </td>
</tr>
<tr>
  <td width="201" height="20">Material Number (Part Number)  <span class="red">*</span> </td>
  <td width="553" height="20" class="red">
      <div align="left">
        <input name="materialnumber" type="text" id="materialnumber" value="<%=rs("MATERIAL_NUMBER")%>">
      </div></td>
    </tr>
  <tr>
    <td height="20">Material Name <span class="red">*</span> </td>
    <td height="20"><input name="materialname" type="text" id="materialname" value="<%=rs("MATERIAL_NAME")%>"></td>
  </tr>
  <tr>
    <td height="20">Belonged Factory <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
        <option value="">-- Select Factory --</option>
        <%= getFactory("OPTION",rs("FACTORY_ID"),factoryiteminside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20">Lot to Be Locked</td>
    <td height="20"><input name="locked" type="text" id="locked" value="<%=rs("LOCKED_LOT")%>" size="50">
      use &quot;,&quot; to seperate lot number. </td>
  </tr>
  <tr>
    <td height="20">Unit <span class="red">*</span> </td>
    <td height="20"><select name="materialunit" id="materialunit">
      <option value="">-- Select --</option>
      <option value="EA" <%if rs("UNIT")="EA" then%>selected<%end if%>>EA</option>
      <option value="ME" <%if rs("UNIT")="ME" then%>selected<%end if%>>ME</option>
      <option value="GM" <%if rs("UNIT")="GM" then%>selected<%end if%>>GM</option>
    </select></td>
  </tr>
  <tr>
    <td height="20">EA Ratio <span class="red">*</span> </td>
    <td height="20">1:
      <input name="ratio" type="text" id="ratio" value="<%=rs("MIN_UNIT_RATIO")%>" size="15"></td>
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
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<%rs.close%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
