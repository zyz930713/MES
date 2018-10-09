<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
pagename="NewEngineer.asp"
path=request.QueryString("path")
query=request.QueryString("query")
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Operator/FormCheck.js" type="text/javascript"></script>
<script language="javascript">
function preload()
{
document.form1.code.focus()
}
</script>
</head>

<body onLoad="preload()">
<form name="form1" method="post" action="/System/Operator/NewOperator1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Create a New Operator </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="t-c-greenCopy">User:
              <% =session("User") %></td>
          <td width="50%"><div align="right"></div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20">Code <span class="red">*</span></td>
      <td height="20"><input name="code" type="text" id="code"></td>
    </tr>
    <tr> 
      <td width="56" height="20">Name <span class="red">*</span></td>
      <td width="707" height="20"><input name="name" type="text" id="name">	  </td>
    </tr>
    <tr>
      <td height="20">Chinese name </td>
      <td height="20"><input name="chinesename" type="text" id="chinesename"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20">Belonged Factory <span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value="">-- Select Section --</option>
          <%= getFactory("OPTION","","","","") %>
      </select></td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
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