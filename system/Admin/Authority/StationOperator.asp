<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStationCodes.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request.QueryString("factory")
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and S.FACTORY_ID is null"
else
where=where&" and S.FACTORY_ID='"&factory&"'"
end if
pagename="/Admin/Authority/StationOperator.asp"
pagepara="&factory="&factory
FactoryRight "S."
SQL="Select 1,S.*,F.FACTORY_NAME from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&factorywhereoutside&" order by F.FACTORY_NAME,S.STATION_NAME"
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
  <td height="20" colspan="7" class="t-c-greenCopy">Browse Operators' Authority (by Station) </td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy">    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><input type="button" name="Button" value="Browse by Operator" onClick="javascript:location.href='Operator.asp'">
          <input type="button" name="Button2" value="Browse by Part" onClick="javascript:location.href='PartOperator.asp'"></td>
        <td><div align="right"><a href="/Admin/Authority/NewOperator.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Operator</a></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td class="t-t-Borrow"><div align="center">Index</div></td>
  <%if admin=true then%>
  <td class="t-t-Borrow"><div align="center">Edit</div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Station Name </div></td>
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
  <td class="t-t-Borrow"><div align="center">Operators</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Count</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
idcount=0
%>
<tr>
  <td><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditStationAuthority.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">Edit</span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("STATION_NAME") %></div></td>
    <td><div align="center"><%= rs("STATION_CHINESE_NAME") %></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME") %></div></td>
    <td><%= getStationCodes(rs("NID"),";",idcount)%>&nbsp;</td>
    <td height="20"><%=idcount%>&nbsp;</td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
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
