<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Material/MaterialCheck.asp" -->
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
where=where&" and M.FACTORY_ID is null"
else
where=where&" and M.FACTORY_ID='"&factory&"'"
end if
pagename="/Admin/Material/Material.asp"
pagepara="&factory="&factory
FactoryRight "M."
SQL="select 1,M.*,F.FACTORY_NAME from MATERIAL M inner join FACTORY F on M.FACTORY_ID=F.NID"&where&factorywhereoutside
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy">Browse Material List </td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Material/AddMaterial.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Material</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td width="45" height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td width="87" class="t-t-Borrow"><div align="center">Status</div></td>
  <td width="120" height="20" class="t-t-Borrow"><div align="center">Material Number </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">Material Name </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">Locked Lot </div></td>
  <td width="196" height="20" class="t-t-Borrow"><div align="center">Unit</div></td>
  <td width="196" class="t-t-Borrow"><div align="center">EA Ratio </div></td>
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
    <td width="44" height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>>
        <%if rs("STATUS")="1" then%>
        <span style="cursor:hand" onClick="javascript:window.open('EditMaterial.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span>
        <%else%>
<img src="/Images/IconEdit.gif" alt="">
        <%end if%>
    </div></td>
    <td width="54" height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>>
        <%if rs("STATUS")="1" then%>
        <span style="cursor:hand" onClick="javascript:window.open('DeleteMaterial.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&materialname=<%=rs("MATERIAL_NAME")%>','main')"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span>
        <%else%>
        <img src="/Images/IconDelete.gif" alt="" width="16" height="20">
        <%end if%>
    </div></td>
    <td class="red"><div align="center">
        <%if rs("STATUS")="1" then%>
        <span style="cursor:hand" onClick="javascript:location.href='DisableMaterial.asp?id=<%=rs("NID")%>&materialnumber=<%=rs("MATERIAL_NUMBER")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Material"><img src="/Images/Enabled.gif"></span>
        <%else%>
        <span style="cursor:hand" onClick="javascript:location.href='EnableMaterial.asp?id=<%=rs("NID")%>&materialnumber=<%=rs("MATERIAL_NUMBER")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Material"><img src="/Images/Disabled.gif"></span>
        <%end if%>
    </div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("MATERIAL_NUMBER") %></div></td>
    <td><div align="center"><%= rs("MATERIAL_NAME") %></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME") %></div></td>
    <td><div align="center"><%= rs("LOCKED_LOT") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("UNIT") %></div></td>
    <td><div align="center"><%= rs("MIN_UNIT_RATIO") %>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="10"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
