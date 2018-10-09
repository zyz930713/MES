<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from PART_MODEL where ITEM_ID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/PartModel/EditPartModel1.asp" method="post" name="form1" target="_self">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Part Model</td>
</tr>
<tr>
  <td width="133" height="20">Modele Number</td>
  <td width="621" height="20"><div align="left"><input name="model_number" type="hidden" id="model_number" value="<%=rs("MODEL_NUMBER")%>"><%=rs("MODEL_NUMBER")%></div></td>
    </tr>
<tr>
  <td height="20">
    Description</td>
  <td height="20"><%=rs("DESCRIPTION")%></td>
</tr>
<tr>
  <td height="20">Factory </td>
  <td height="20"><%=rs("WIP_SUPPLY_SUBINVENTORY")%></td>
</tr>
<tr>
  <td height="20">Lead Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
      <input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
      <select name="lead_time_unit" id="lead_time_unit">
        <option value="">-- Select--</option>
        <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
        <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
        <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
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
