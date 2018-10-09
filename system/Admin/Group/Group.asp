<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Group/GroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.QueryString("factory")
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and SG.FACTORY_ID is null"
else
where=where&" and SG.FACTORY_ID='"&factory&"'"
end if
pagename="/Admin/Group/Group.asp"
pagepara="&factory="&factory
FactoryRight "SG."
SQL="select 1,SG.*,F.FACTORY_NAME from SYSTEM_GROUP SG inner join FACTORY F on SG.FACTORY_ID=F.NID "&factorywhereoutside&" order by SG.NID"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page();language(<%=session("language")%>);">

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
        <a href="/Admin/Group/AddGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span> </a>
        <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="11"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_SearchStatus"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Name"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_CHName"></span></div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_GroupType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Member"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Total"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
  <tr <%if rs("STATUS")="0" then%>class="t-b-blue"<%end if%>>
    <td height="20"><div align="center">
      <% =(cint(session("strpagenum"))-1)*recordsize+i%>
    </div></td>
	<%if admin=true then%>
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:window.open('EditGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit" width="16" height="20" border="0" align="absmiddle"></span><%end if%>
	&nbsp;</div></td>
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('DeleteGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&groupname=<%=rs("GROUP_NAME")%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20" border="0" align="absmiddle"></span><%end if%>
	&nbsp;</div></td>
    <td class="red"><div align="center"><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:location.href='DisableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Group"><img src="/Images/Enabled.gif"></span><%else%><span style="cursor:hand" onClick="javascript:location.href='EnableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Group"><img src="/Images/Disabled.gif"></span><%end if%></div></td>
	<%end if%>
<td height="20"><div align="center"><%=rs("GROUP_NAME")%></div></td>
    <td><div align="center"><%=rs("GROUP_CHINESE_NAME")%></div></td>
    <td><div align="center"><%=rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%=rs("GROUP_TYPE")%></div></td>
    <td><div align="left">
	<%if rs("GROUP_MEMBERS")<>"" then
	amembers=split(rs("GROUP_MEMBERS"),",")	
	for j=1 to ubound(amembers)+1 
		md=j mod 8
		if md<>0 then%>
		<%=amembers(j-1)&","%>
	<%	else%>
		<%=amembers(j-1)&",<br>"%>
	<%	end if
	next%>
    </div></td>
    <td><div align="center"><%=ubound(amembers)+1%></div></td>
    <%else%>
        &nbsp;</div></td>
        <td><div align="center">0</div></td>
    <%end if%>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="11"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->