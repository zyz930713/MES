<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetDefectCode.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
group_name=request("group_name")
factory=request("factory")
if defectcode_group_name<>"" then
where=where&" and lower(DG.GROUP_NAME) like '%"&lcase(group_name)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and DG.STATUS="&thisstatus
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and DG.FACTORY_ID is null"
else
where=where&" and DG.FACTORY_ID='"&factory&"'"
end if

pagename="/Admin/DefectCodeGroup/DefectCodeGroup.asp"
pagepara="&group_name="&group_name&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "DG."
SQL="select DG.*,F.FACTORY_NAME from DEFECTCODE_GROUP DG inner join FACTORY F on DG.FACTORY_ID=F.NID where 1=1 "&where&factorywhereoutsideand&" order by DG.GROUP_NAME"
session("SQL")=SQL
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
<form action="/Admin/DefectCodeGroup/DefectCodeGroup.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search DefectCode Group</td>
  </tr>
  <tr>
    <td height="20">DefectCode Group Name </td>
    <td><input name="defectcode_group_name" type="text" id="defectcode_group_name" value="<%=defectcode_group_name%>"></td>
    <td>Factory</td>
    <td><select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy">Browse DefectCode Group List</td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/DefectCodeGroup/AddDefectCodeGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New DefectCode Group</a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">DefectCode Group Name </div></td>
  <td class="t-t-Borrow"><div align="center">Chinese Name</div></td>
  <td class="t-t-Borrow"><div align="center">Factory</div></td>
  <td class="t-t-Borrow"><div align="center">Included Defectcodes </div></td>
  </tr>
  <form name="checkform" method="post" action="/Admin/DefectCodeGroup/UpdateDefectCodeGroupAll.asp">
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
  <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditDefectCodeGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Series Group?')){window.open('DeleteDefectCodeGroup.asp?id=<%=rs("NID")%>&defectcode_group_name=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("GROUP_NAME") %></div></td>
    <td><div align="center"><%= rs("GROUP_CHINESE_NAME") %></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%if rs("MEMBERS_ID")<>"" then%><%=getDefectcode("TEXT",null," where D.NID in ('"&replace(rs("MEMBERS_ID"),",","','")&"')",null," ; ")%><%end if%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
  
<%
else
%>
</form>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->