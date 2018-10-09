<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
pagename="/System/Role/Role.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
thisname=trim(request("thisname"))
if ordername="" and ordertype="" then
order=" order by ROLE_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
if thisname<>"" then
where=where&" and lower(ROLE_NAME) like '%"&lcase(thisname)&"%'"
end if
pagepara="&thisname="&thisname
pagename="/System/Role/Role.asp" 
SQL="select * from ENGINEER_ROLE where NID is not null "&where&order
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form name="form1" method="post" action="/System/Role/Role.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_Name"></span></td>
    <td ><input name="thisname" type="text" id="thisname" value="<%=thisname%>"></td>
    <td ><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="white"><span id="inner_User"></span>:<%=session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Role/NewRole.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a> </div></td>
      </tr>
    </table>   </td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td width="2%" height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td width="4%" colspan="2"><div align="center"><span id="inner_Action"></span></div></td>
    <td width="5%"><div align="center"><span id="td_Apply"></span></div></td>
    <td width="5%" height="20"><div align="center"><span id="td_Name"></span></div></td>
    <td width="10%"><div align="center"><span id="td_CHName"></span></div></td>
    <td width="10%"><div align="center"><span id="td_Description"></span></div></td>
    <td width="11%"><div align="center"><span id="td_Applicant"></span></div></td>
    <td width="11%"><div align="center"><span id="td_Member"></span></div></td>
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*pagesize_s+i%>
    </div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('/System/Role/EditRole.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this role? If deleted, engineer will be excluded this role.\n您确定删除此角色吗？如果删除，工程师对应的角色将删除。')){window.open('/System/Role/DeleteRole.asp?id=<%=rs("NID")%>&rolename=<%=rs("ROLE_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif"></span></div></td>
    <td><div align="center"><%if rs("APPLY_FROM_WEB")="1" then%><img src="/Images/Yes.gif" width="16" height="16"><%end if%>&nbsp;</div></td>
    <td height="20"><div align="left"><%=rs("ROLE_NAME")%></div></td>
    <td><div align="center"><%=rs("ROLE_CHINESE_NAME")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("DESCRIPTION")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("APPLICANT")%>&nbsp;</div></td>
    <td><div align="center"><%=getRoleMember(true,"TEXT",""," where ROLES_ID like '%"&rs("ROLE_NAME")&"%'","",""," ; ")%>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->