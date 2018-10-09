<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/OracleLine/OracleLineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
line_id=request.QueryString("line_id")
SQL="select L.LINE_CODE,L.ORGANIZATION_ID,S.SubQuantity from tbl_Wip_Lines L left join tbl_Wip_Line_Sub S on L.Line_Id=S.Line_Id where L.LINE_ID='"&line_id&"'"
rsPR.open SQL,connPR,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript">
</script>
<script language="JavaScript" src="/Admin/OracleLine/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/OracleLine/EditOracleLine1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit an Oracle Line   </td>
</tr>
<tr>
  <td width="123" height="20"><div align="left">Line Name <span class="red">*</span> </div></td>
    <td width="631" height="20">
      <div align="left">
        <%=rsPR("LINE_CODE")%>
      </div></td>
    </tr>
<tr>
  <td height="20">Organization <span class="red">*</span></td>
  <td height="20"><%= rsPR("ORGANIZATION_ID")%></td>
</tr>
<tr>
  <td height="20">Sublot Quantity </td>
  <td height="20"><input name="subqty" type="text" id="subqty" value="<%=rsPR("SubQuantity")%>"></td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="line_id" type="hidden" id="line_id" value="<%=line_id%>">
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
<!--#include virtual="/WOCF/POCF_Close.asp" -->