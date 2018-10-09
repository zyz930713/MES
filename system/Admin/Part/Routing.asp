<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%

partnumber=request("partnumber")
partrule=request("partrule")
factory=request.QueryString("factory")
where=""
if partnumber<>"" then
where=where&" and lower(R.PART_NUMBER) like '%"&lcase(partnumber)&"%'"
end if
if partrule<>"" then
where=where&" and lower(R.PART_RULE) like '%"&lcase(partrule)&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and R.FACTORY_ID is null"
else
where=where&" and R.FACTORY_ID='"&factory&"'"
end if
pagepara="&partnumber="&partnumber&"&partrule="&partrule&"&factory="&factory
FactoryRight "R."
SQL="select 1,R.*,S.SECTION_NAME,F.FACTORY_NAME from ROUTING R left join SECTION S on R.SECTION_ID=S.NID inner join FACTORY F on R.FACTORY_ID=F.NID where R.IS_DELETE=0 "&where&factorywhereoutsideand&" order by R.NID"
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
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language(<%=session("language")%>);language_page()">
<form action="/Admin/Part/Routing.asp" method="get" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_RoutingNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td height="20"><span id="inner_RoutingRule"></span></td>
    <td><input name="partrule" type="text" id="partrule" value="<%=partrule%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="15" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="15" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><a href="/Admin/Part/AddRouting_New.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddData"></span></a></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="15"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="3" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_RoutingNumber"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_RoutingRule"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_RoutingType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Priority"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LineName"></span></div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Section"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_StationSeq"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_TargetYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_MaxInterval"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
if rs("ROUTINE_TYPE")="0" then
main=true
else
main=false
end if
%>
<tr<%if main=false then%> <%end if%>>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
	<td height="20" ><div align="center">	
	 <span style="cursor:hand" onClick="javascript:window.open('CopyRouting.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
		<img src="/Images/IconCopy.gif" alt="Click to copy routing"></span></div>
	 </td>
	 <td height="20" ><div align="center">
	 <span class="red" style="cursor:hand" onClick="javascript:window.open('EditRouting.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
		<img src="/Images/IconEdit.gif" alt="Click to edit"></span></div>
	</td>    
	<td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this part? If deleted, new job can not apply it forever. ')){window.open('DeleteRouting.asp?id=<%=rs("NID")%>&partnumber=<%=rs("PART_NUMBER")%>&path=<%=path%>&query=<%=query%>','main')}">
		<img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div>
	</td> 
  <%end if%>	
	<td height="20"><div align="center"><%= rs("PART_NUMBER") %></div></td>
	<td nowrap><div align="center"><%= replace(rs("PART_RULE"),",",",<br>") %></div></td>
	 <td><div align="center">
	 <%if rs("PART_TYPE")="0" then %>Normal
	 <%elseif rs("PART_TYPE")="1" then%>Rework
	 <%elseif rs("PART_TYPE")="4" then%>Retest
	 <%elseif rs("PART_TYPE")="5" then%>Slapping
	 <%end if%>
	 
	 </div></td>
    <td><div align="center"><%=rs("MEET_PRIORITY")%></div></td>
    <td><div align="center">
        <%if not isnull(rs("LINES_INDEX")) or rs("LINES_INDEX")<>"" then%>
        <%= getLine("TEXT",""," where L.NID in ('"&replace(rs("LINES_INDEX"),",","','")&"')","","<br>") %>
                   <%else%>&nbsp;
        <%end if%>
    </div></td>
    <td><div align="center"><%= rs("FACTORY_NAME") %></div></td>
    <td><div align="center"><%= rs("SECTION_NAME") %>&nbsp;</div></td>
    <td height="20"><div align="left"><%if rs("STATIONS_INDEX")<>"" then%><%= getStation_New2(true,"TEXT_LINK",""," and S.NID in ('"&replace(rs("STATIONS_INDEX"),",","','")&"')","",rs("STATIONS_INDEX")," -> ") %><%end if%>&nbsp;</div></td>
    <td><div align="center" title="<%= rs("STATIONS_TRANSACTION") %>">
        <%if rs("STATIONS_ROUTINE")="0" then%>
        Fixed
        <%else%>
        Repeated
        <%end if%>
    </div></td>
    <td><div align="center"><%= rs("TARGET_YIELD")%>%</div></td>
    <td><div align="center"><%= rs("MAX_INTERVAL") %>&nbsp;</div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="15" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->