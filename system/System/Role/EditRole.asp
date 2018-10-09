<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="NewRole.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from ENGINEER_ROLE where NID='"&id&"'"
rs.open SQL,conn,1,3
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
<form name="form1" method="post" action="/System/Role/EditRole1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="138" height="20"><span id="td_Name"></span> <span class="red">*</span></td>
      <td  height="20"><input name="name" type="text" id="name" value="<%=rs("ROLE_NAME")%>" size="50">	  </td>
    </tr>
    <tr>
      <td height="20"><span id="td_CHName"></span> <span class="red">*</span> </td>
      <td height="20"><input name="chinesename" type="text" id="chinesename" value="<%=rs("ROLE_CHINESE_NAME")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Description"></span></td>
      <td height="20"><input name="description" type="text" id="description" value="<%=rs("DESCRIPTION")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="20"><span id="td_Applicant"></span></td>
      <td height="20"><input name="applicant" type="text" id="applicant" value="<%=rs("APPLICANT")%>" size="100"></td>
    </tr>
    <tr>
      <td height="20"><span id="td_Member"></span></td>
      <td height="20"><%old_memebers=getRoleMember(false,"CODE",role," where ROLES_ID like '%"&rs("ROLE_NAME")&"%'"," order by USER_CODE",null,",")%>
	  <%=getRoleMember(false,"CHECK",role," where ROLES_ID like '%"&rs("ROLE_NAME")&"%'"," order by USER_CODE",null,null)%></td>
    </tr>
    <tr>
      <td height="20"><span id="td_ApplyDesc"></span></td>
      <td height="20"><input name="apply" type="checkbox" id="apply" value="1"<%if rs("APPLY_FROM_WEB") then%> checked<%end if%>></td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input name="old_memebers" type="hidden" id="old_memebers" value="<%=old_memebers%>">
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