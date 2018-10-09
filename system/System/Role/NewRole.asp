<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="NewRole.asp"
path=request.QueryString("path")
query=request.QueryString("query")
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Role/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onload="language(<%=session("language")%>);">
<form name="form1" method="post" action="/System/Role/NewRole1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="106" height="20"><span id="td_Name"></span> <span class="red">*</span></td>
      <td width="913" height="20"><input name="name" type="text" id="name" size="50">	  </td>
    </tr>
    <tr>
      <td height="20"><span id="td_CHName"></span> <span class="red">*</span> </td>
      <td height="20"><input name="chinesename" type="text" id="chinesename" size="50"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Description"></span></td>
      <td height="20"><input name="description" type="text" id="description" size="50"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Applicant"></td>
      <td height="20"><input name="applicant" type="text" id="applicant" size="100"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_ApplyDesc"></span></td>
      <td height="20"><input name="apply" type="checkbox" id="apply" value="1"></td>
    </tr>
    
    
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="btnOK" value="OK">
          &nbsp; 
          <input type="reset" name="btnReset" value="Reset">
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->