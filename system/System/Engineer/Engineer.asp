<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetEvent.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request("factory")
code=trim(request("code"))
thisname=trim(request("thisname"))
account=request("account")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by U.USER_CODE asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if factory="" or factory="all" then
where=where&""
else
where=where&" and U.FACTORY_ID='"&factory&"'"
end if
if code<>"" then
where=where&" and U.USER_CODE like '%"&code&"%'"
end if
if thisname<>"" then
where=where&" and lower(U.USER_NAME) like '%"&lcase(thisname)&"%'"
end if
if account<>"" then
where=where&" and lower(U.NT_ACCOUNT) like '%"&lcase(account)&"%'"
end if
pagepara="&factory="&factory&"&code="&code&"&thisname="&thisname&"&account="&account
pagename="/System/Engineer/Engineer.asp" 
SQL="select U.*,UM.USER_NAME as MANAGER_NAME from USERS U left join USERS UM on U.MANAGER=UM.USER_CODE where U.NID is not null "&where&order
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form name="form1" method="post" action="/System/Engineer/Engineer.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_Code"></span></td>
    <td height="20"><input name="code" type="text" id="code" value="<%=code%>"></td>
    <td><span id="inner_Name"></span></td>
    <td><input name="thisname" type="text" id="thisname" value="<%=thisname%>"></td>
    <td><span id="inner_NTAccount"></span></td>
    <td><input name="account" type="text" id="account" value="<% =account%>"></td>
    <td><span id="inner_SearchFactory"></span></td>
    <td><select name="factory" id="factory">
      <option value="">All</option>      
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
          <% =session("User") %>        </td>
        <td width="50%"><div align="right"><a href="/System/Engineer/NewEngineer.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20" colspan="2"><div align="center"><span id="inner_Action"></span></div></td>
    <td height="20"><div align="center"><span id="inner_SearchStatus"></span></div></td>
    <td height="20"><div align="center"><span id="td_Code"></span></div></td>
    <td height="20"><div align="center"><span id="td_Name"></span></div></td>
    <td><div align="center"><span id="td_CHName"></span></div></td>
    <td height="20"><div align="center"><span id="td_NTAccount"></span></div></td>
    <td><div align="center"><span id="td_Manager"></span></div></td>
    <td><div align="center"><span id="td_Factory"></span></div></td>
    <td><div align="center"><span id="td_Language"></span></div></td>
    <td height="20"><div align="center"><span id="td_Role"></span></div></td>
    <td height="20"><div align="center"><span id="td_Email"></span></div></td>
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
        <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('/System/Engineer/EditEngineer.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('/System/Engineer/DeleteEngineer.asp?id=<%=rs("NID")%>&username=<%=rs("USER_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif"></span></div></td>
    <td height="20"><div align="center">
        <%if rs("status")="1" then%>
        <span class="red" style="cursor:hand"title="Click to Disable" onClick="javascript:window.open('DisableEngineer.asp?id=<%=rs("NID")%>&rolename=<%=rs("USER_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/Enabled.gif"></span>
        <%else%>
        <span class="red" style="cursor:hand" title="Click to Enable" onClick="javascript:window.open('EnableEngineer.asp?id=<%=rs("NID")%>&rolename=<%=rs("USER_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/Disabled.gif"></span>
        <%end if%>
    </div></td>
    <td height="20"><div align="center"><%=rs("USER_CODE")%></div></td>    
    <td height="20"><div align="center"><%=rs("USER_NAME")%></div></td>
    <td><div align="center"><%=rs("USER_CHINESE_NAME")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("NT_ACCOUNT")%></div></td>
    <td><div align="center"><%=rs("MANAGER_NAME")%>&nbsp;</div></td>
    <td><div align="center"><%=getFactory("TEXT",null," where instr('"&rs("FACTORY_ID")&"',NID)>0",null,"; ")%>&nbsp;</div></td>
    <td><div align="center"><%if cint(rs("LANGUAGE"))=0 then%>ENG<%else%>CHN<%end if%></div></td>
    <td height="20"><div align="left"><%if rs("ROLES_ID")<>"" then%><%=replace(rs("ROLES_ID"),",","<br>")%><%end if%>&nbsp;</div></td>
    <td height="20"><%=rs("EMAIL")%>&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center"><span id="inner_Records"></span>&nbsp;</div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->