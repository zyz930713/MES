<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<%
pagename="NewEngineer.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from OPERATORS where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Operator/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form name="form1" method="post" action="/System/Operator/EditOperator1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Edit a  Operator </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="white">User:
              <% =session("User") %></td>
          <td width="50%"><div align="right"></div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20">Code <span class="red">*</span></td>
      <td height="20"><input name="code" type="text" id="code" value="<%=rs("CODE")%>"></td>
    </tr>
    <tr> 
      <td width="56" height="20">Name <span class="red">*</span></td>
      <td width="707" height="20"><input name="name" type="text" id="name" value="<%=rs("OPERATOR_NAME")%>">
	  </td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="Submit" value="Submit">
          &nbsp; 
          <input type="reset" name="Submit2" value="Reset">
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->