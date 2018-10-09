<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/GetDefectCode.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
defectcodename=request.QueryString("defectcodename")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="post" action="">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy">System adjustmenet before delete defectocode </td>
    </tr>
  <tr>
    <td width="14%" height="20">Deleteed Defectcode </td>
    <td width="86%" height="20"><%= defectcodename %>&nbsp;</td>
  </tr>
  <tr>
    <td height="20">Replaced Defectcode</td>
    <td height="20">
      <select name="replace_id" id="replace_id">
	  <option value="">-- Select DefectCode--</option>
	  <%=getDefectCode("where NID<>'"&id&"'")%>
      </select>    </td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input type="submit" name="Submit" value="Submit">
      &nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->