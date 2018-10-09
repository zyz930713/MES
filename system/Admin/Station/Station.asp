<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
stationname=request("stationname")
chinesestationname=request("chinesestationname")
factory=request.QueryString("factory")
stationNID=request("stationNID")
if stationNID<>"" then
	where=where&" and ST.NID='"&stationNID&"'"
end if
if stationname<>"" then
where=where&" and lower(ST.STATION_NAME) like '%"&lcase(stationname)&"%'"
end if
if chinesestationname<>"" then
where=where&" and lower(ST.STATION_CHINESE_NAME) like '%"&lcase(chinesestationname)&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and S.FACTORY_ID is null"
else
where=where&" and S.FACTORY_ID='"&factory&"'"
end if
pagename="/Admin/Station/Station.asp"
pagepara="&factory="&factory
FactoryRight "ST."
SQL="select 1,ST.*,S.SECTION_NAME,F.FACTORY_NAME from STATION ST left join SECTION S on ST.SECTION_ID=S.NID left join FACTORY F on ST.FACTORY_ID=F.NID where ST.STATUS=1 "&where&factorywhereoutsideand&"order by ST.STATION_NUMBER"
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
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language_page()">
<form action="/Admin/Station/Station.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn">Search Station</td>
  </tr>
  <tr>
  	 <td height="20">Station ID</td>
    <td><input name="stationNID" type="text" id="stationNID" value="<%=stationNID%>"></td>
    <td height="20">Station Name</td>
    <td><input name="stationname" type="text" id="stationname" value="<%=stationname%>"></td>
    <td height="20">Station Chinese Name</td>
    <td><input name="chinesestationname" type="text" id="chinesestationname" value="<%=chinesestationname%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">Browse Station list </td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Station/AddStation.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Station</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="1" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center">NID</div></td>
  <td class="t-t-Borrow"><div align="center">Station Number</div></td>
  <td class="t-t-Borrow"><div align="center">Station Name</div></td>
  <td class="t-t-Borrow"><div align="center">Chinese Name </div></td>
  <td class="t-t-Borrow"><div align="center">
    <div align="center">
      <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
        <option value="">Factory</option>
        <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
        <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
      </select>
    </div>
  </div></td>
  <td class="t-t-Borrow"><div align="center">Section</div></td>
  <td class="t-t-Borrow"><div align="center">WIP</div></td>
  <td class="t-t-Borrow"><div align="center">WIP SEQ </div></td>
  <td class="t-t-Borrow"><div align="center">Output</div></td>
  <td class="t-t-Borrow"><div align="center">Output SEQ </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Included Actions</div></td>
  <td class="t-t-Borrow"><div align="center">Stations with Defect Code</div></td>
  <td class="t-t-Borrow"><div align="center">Ini</div></td>
  <td class="t-t-Borrow"><div align="center">WIP Merged Stations</div></td>
</tr>
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
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditStation.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <!--<td height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Station? If deleted, new job can not apply it forever. ')){window.open('DeleteStation.asp?id=<%'=rs("NID")%>&stationname=<%'=rs("STATION_NAME")%>&path=<%'=path%>&query=<%'=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>-->
	<%end if%>
	<td><div align="center"><%= rs("NID") %></div></td>
    <td><div align="center"><%= rs("STATION_NUMBER") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%if rs("TRANSACTION_TYPE")="1" then%><span class="italic"><%= rs("STATION_NAME") %></span><%else%><%= rs("STATION_NAME") %><%end if%></div></td>
    <td><div align="center"><%= rs("STATION_CHINESE_NAME") %></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME") %></div></td>
    <td><div align="center"><%= rs("SECTION_NAME") %>&nbsp;</div></td>
    <td>      <div align="center">
        <%if rs("WIP_REPORT_COLUMN")="1" then%>
      Y
      <%end if%>
&nbsp;    </div></td>
    <td><div align="center"><%= rs("WIP_SEQUENCY") %>&nbsp;</div></td>
    <td><div align="center">
        <%if rs("OUTPUT_REPORT_COLUMN")="1" then%>
      Y
      <%end if%>
      &nbsp; </div></td>
    <td><div align="center"><%= rs("OUTPUT_SEQUENCY") %>&nbsp;</div></td>
    <td height="20"><div align="left"><%if not isnull(rs("ACTIONS_INDEX")) then%><%= getAction(true,"TEXT",""," where A.NID in ('"&replace(rs("ACTIONS_INDEX"),",","','")&"')","",rs("ACTIONS_INDEX")," ; ") %><%end if%>&nbsp;</div></td>
    <td><div align="center">
        <%if rs("STATION_ENTER_DEFECTCODE")<>"" then%>
        <%= getStation(true,"TEXT",""," where S.NID in ('"&replace(rs("STATION_ENTER_DEFECTCODE"),",","','")&"')","",rs("STATION_ENTER_DEFECTCODE")," ; ") %>
        <%end if%>
    &nbsp;</div></td>
    <td><%= rs("INITAIL_QUANTITY_TYPE") %></td>
    <td><%if rs("WIP_INCLUDED_STATIONS")<>"" then%>
      <%= getStation(true,"TEXT",""," where S.NID in ('"&replace(rs("WIP_INCLUDED_STATIONS"),",","','")&"')","",rs("WIP_INCLUDED_STATIONS")," -> ") %>
      <%end if%>
&nbsp; </td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>  
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->