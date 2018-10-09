<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="NewEvent.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from EVENT where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Event/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form name="form1" method="post" action="/System/Event/EditEvent1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Edit a Event</td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">User:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="106" height="20">Name <span class="red">*</span></td>
      <td width="913" height="20"><input name="name" type="text" id="name" value="<%=rs("EVENT_NAME")%>" size="50">	  </td>
    </tr>
    <tr> 
      <td height="20">Description</td>
      <td height="20"><input name="description" type="text" id="description" value="<%=rs("DESCRIPTION")%>">      </td>
    </tr>
    <tr>
      <td height="20">Note</td>
      <td height="20"><textarea name="note" cols="50" rows="5" id="note"><%=rs("EVENT_NOTE")%></textarea></td>
    </tr>
    <tr>
      <td height="20">Email body</td>
      <td height="20"><textarea name="body" cols="50" rows="5" id="body"><%=rs("EMAIL_BODY")%></textarea></td>
    </tr>
    <tr>
      <td height="20">Email Hyperlink </td>
      <td height="20"><input name="hyperlink" type="text" id="hyperlink" value="<%=rs("EMAIL_HYPERLINK")%>" size="100"></td>
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