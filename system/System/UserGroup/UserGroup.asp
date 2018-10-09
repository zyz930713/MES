<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetUser.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.QueryString("factory")
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and UG.FACTORY_ID is null"
else
where=where&" and UG.FACTORY_ID='"&factory&"'"
end if
pagename="/System/UserGroup/UserGroup.asp"
pagepara="&factory="&factory
FactoryRight "UG."
SQL="select 1,UG.*,F.FACTORY_NAME from USER_GROUP UG inner join FACTORY F on UG.FACTORY_ID=F.NID "&factorywhereoutside&" order by UG.NID"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">Browse User Group List </td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right">
        <a href="/System/UserGroup/AddUserGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New User Group </a>
        </div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <td class="t-t-Borrow"><div align="center">Status</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Group Name </div></td>
  <td class="t-t-Borrow"><div align="center">Chinese Name </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">Role</div></td>
  <td class="t-t-Borrow"><div align="center">Members</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr <%if rs("STATUS")="0" then%>class="t-b-blue"<%end if%>>
    <td height="20"><div align="center">
      <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:window.open('EditUserGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit" width="16" height="20" border="0" align="absmiddle"></span><%end if%>
</div></td>
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this group? If so, all sepcial settings in Action for this group will be deleted. ')){window.open('DeleteUserGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&groupname=<%=rs("GROUP_NAME")%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20" border="0" align="absmiddle"></span><%end if%></div></td>
    <td class="red"><div align="center"><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:location.href='DisableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Group"><img src="/Images/Enabled.gif"></span><%else%><span style="cursor:hand" onClick="javascript:location.href='EnableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Group"><img src="/Images/Disabled.gif"></span><%end if%></div></td>
<td height="20"><div align="center"><%=rs("GROUP_NAME")%></div></td>
    <td><div align="center"><%=rs("GROUP_CHINESE_NAME")%></div></td>
    <td><div align="center"><%=rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%=formatlongstringbreak(rs("ROLES_ID"),"<br>",30)%>&nbsp;</div></td>
    <td><div align="left"><%if rs("GROUP_MEMBERS")<>"" then%><%=getUser(true,"TEXT",""," where USER_CODE in ('"&replace(rs("GROUP_MEMBERS"),",","','")&"')","",""," ; ")%><%end if%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->