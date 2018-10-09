<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Group/GroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from SYSTEM_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Group/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">

<form action="/Admin/Group/EditGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
</tr>
<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
<tr>
  <td width="135" height="20"><div align="left"><span id="td_Name"></span> <span class="red">*</span> </div></td>
  <td width="533" height="20" class="red">
    <div align="left">
      <input name="groupname" type="text" id="groupname" value="<%=rs("GROUP_NAME")%>" size="50">
  </div></td>
</tr>
<tr>
  <td height="20"><span id="td_CHName"></span> <span class="red">*</span></td>
  <td height="20"><input name="groupchinesename" type="text" id="groupchinesename" value="<%=rs("GROUP_CHINESE_NAME")%>" size="50"></td>
</tr>
<tr>
  <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value=""></option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><div align="left"><span id="td_GroupType"></span> <span class="red">*</span> </div></td>
  <td height="20"><%=rs("GROUP_TYPE") %>
    <input type="hidden" name="grouptype" id="grouptype" value="<%=rs("GROUP_TYPE") %>">  
  </td>
</tr>
<tr>
  <td height="20"><span id="td_Member"></span></td>
  <td height="20"><textarea name="members" cols="100" rows="5" id="members"><%=rs("GROUP_MEMBERS")%></textarea> 
    <br>
    use &quot;,&quot; to separate members. 使用&quot;,&quot;分隔不同成员。 </td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
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
