<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Machine/MachineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
machinenumber=request("machinenumber")
machinename=request("machinename")
factory=request.QueryString("factory")
if machinenumber<>"" then
where=where&" and M.MACHINE_NUMBER like '%"&machinenumber&"%'"
end if
if machinename<>"" then
where=where&" and M.MACHINE_NAME like '%"&machinename&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and M.FACTORY_ID is null"
else
where=where&" and M.FACTORY_ID='"&factory&"'"
end if
pagename="/Admin/Machine/Machine.asp"
pagepara="&factory="&factory
FactoryRight "M."
SQL="select 1,M.*,F.FACTORY_NAME from MACHINE M inner join FACTORY F on M.FACTORY_ID=F.NID "&where&factorywhereoutside&"order by M.NID"
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
<form action="/Admin/Machine/Machine.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search Machine </td>
  </tr>
  <tr>
    <td>Machine Number </td>
    <td><input name="machinenumber" type="text" id="machinenumber" value="<%=machinenumber%>"></td>
    <td height="20">Machine Name</td>
    <td><input name="machinename" type="text" id="machinename" value="<%=machinename%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Machine List </td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Machine/AddMachine.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Machine</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <td class="t-t-Borrow"><div align="center">Status</div></td>
  <td class="t-t-Borrow"><div align="center">Locked</div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Machine Number </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Machine Name </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">Stations to be used </div></td>
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
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:window.open('EditMachine.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span>
        <%end if%>
</div></td>
    <td height="20"><div align="center" <%if rs("STATUS")="1" then%>class="red"<%end if%>><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:window.open('DeleteMachine.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&machinename=<%=rs("MACHINE_NUMBER")%>','main')"><img src="/Images/IconDelete.gif" alt="Click to delete"></span>
        <%end if%></div></td>
    <td class="red"><div align="center"><%if rs("STATUS")="1" then%><span style="cursor:hand" onClick="javascript:location.href='DisableMachine.asp?id=<%=rs("NID")%>&machinenumber=<%=rs("MACHINE_NUMBER")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Machine"><img src="/Images/Enabled.gif" width="10" height="10"></span>
        <%else%><span style="cursor:hand" onClick="javascript:location.href='EnableMachine.asp?id=<%=rs("NID")%>&machinenumber=<%=rs("MACHINE_NUMBER")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Machine"><img src="/Images/Disabled.gif"></span><%end if%></div></td>
		<%end if%>
<td><div align="center">
  <%if rs("LOCKED")="1" then%>
  Locked
  <%else%>
  Unlocked
  <%end if%>
</div></td>
    <td height="20"><div align="center"><%=rs("MACHINE_NUMBER")%></div></td>
    <td height="20"><div align="center"><%=rs("MACHINE_NAME")%></div></td>
    <td><div align="center"><%=rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%if rs("STATIONS_USED")<>"" then%>
      <%= getStation(true,"TEXT",""," where S.NID in ('"&replace(rs("STATIONS_USED"),",","','")&"')","",rs("STATIONS_USED")," ; ") %>
      <%end if%>
      &nbsp;
    &nbsp;</div></td>
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