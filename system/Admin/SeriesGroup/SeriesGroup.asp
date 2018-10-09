<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
seriesgroupname=request("seriesgroupname")
thisstatus=request("status")
factory=request("factory")
if seriesgroupname<>"" then
where=where&" and lower(SG.SERIES_GROUP_NAME) like '%"&lcase(seriesgroupname)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and SG.STATUS="&thisstatus
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and SG.FACTORY_ID is null"
else
where=where&" and SG.FACTORY_ID='"&factory&"'"
end if

pagename="/Admin/SeriesGroup/SeriesGroup.asp"
pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "SG."
SQL="select SG.*,F.FACTORY_NAME,S.SECTION_NAME from SERIES_GROUP SG inner join FACTORY F on SG.FACTORY_ID=F.NID left join SECTION S on SG.SECTION_ID=S.NID where 1=1 "&where&factorywhereoutsideand&" order by SG.SERIES_GROUP_NAME"
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
<form action="/Admin/SeriesGroup/SeriesGroup.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn">Search Series Group</td>
  </tr>
  <tr>
    <td height="20">Series Group Name </td>
    <td><input name="seriesgroupname" type="text" id="seriesgroupname" value="<%=seriesgroupname%>"></td>
    <td>Status</td>
    <td><span class="t-t-Borrow">
      <select name="status" id="status" onChange="location.href='<%=pagename%>?status='+this.options[this.selectedIndex].value+'&station=<%=station%>&factory=<%=factory%>&purpose=<%=purpose%>'">
        <option value="">Status</option>
        <option value="all" <%if thisstatus="all" then%>selected<%end if%>>All</option>
        <option value="1" <%if thisstatus="1" then%>selected<%end if%>>Enabled</option>
        <option value="0" <%if thisstatus="0" then%>selected<%end if%>>Disabled</option>
      </select>
    </span></td>
    <td>Facotry</td>
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
    <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">Browse Series Group List </td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Print List" onClick="javascript:window.open('SeriesGroupPrint.asp')"><img src="/Images/Print.gif" width="30" height="30"></span><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('SeriesGroup_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right">
          <%'if admin=true then%>
          <a href="/Admin/SeriesGroup/AddSeriesGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Series Group</a>
          <%'end if%>
        </div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="18"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
    <%if capacity=true then%>
    <td height="20" class="t-t-Borrow"><div align="center">Capacity</div></td>
    <%end if%>
    <%if admin=true then%>
    <td height="20" class="t-t-Borrow"><div align="center">Select</div></td>
    <td height="20" colspan="3" class="t-t-Borrow"><div align="center">Action</div></td>
    <%end if%>
    <td height="20" class="t-t-Borrow"><div align="center">Series Group Name </div></td>
    <td class="t-t-Borrow"><div align="center">Factory</div></td>
    <td class="t-t-Borrow"><div align="center">Section</div></td>
    <td class="t-t-Borrow"><div align="center">Internal Yield </div></td>
    <td class="t-t-Borrow"><div align="center">Target Yield </div></td>
    <td class="t-t-Borrow"><div align="center">Exceptiom</div></td>
    <td class="t-t-Borrow"><div align="center">Lead Time </div></td>
    <td class="t-t-Borrow"><div align="center">WIP Time </div></td>
    <td class="t-t-Borrow"><div align="center">Capacity</div></td>
    <td class="t-t-Borrow"><div align="center">Included Seires </div></td>
    <td class="t-t-Borrow"><div align="center">Included Models </div></td>
  </tr>
  <form name="checkform" method="post" action="/Admin/SeriesGroup/UpdateSeriesGroupAll.asp">
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
      <%if capacity=true then%>
      <td><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('CapacitySeriesGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">C</span></div></td>
      <%end if%>
      <%if admin=true then%>
      <td><div align="center">
        <input name="nid<%=i%>" type="hidden" id="nid<%=i%>" value="<%=rs("NID")%>">
        <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">
      </div></td>
      <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditSeriesGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
      <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Series Group?')){window.open('DeleteSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_GROUP_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
      <td><div align="center"><span class="red">
        <%if rs("STATUS")="1" then%>
        <span style="cursor:hand" onClick="javascript:location.href='DisableSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Series Group"><img src="/Images/Enabled.gif"></span>
        <%else%>
        <span style="cursor:hand" onClick="javascript:location.href='EnableSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Series Group"><img src="/Images/Disabled.gif"></span>
        <%end if%>
      </span></div></td>
      <%end if%>
      <td height="20"><div align="center"><%= rs("SERIES_GROUP_NAME") %></div></td>
      <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
      <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
      <td><div align="center"><%= formatpercent(csng(rs("TARGET_INTERNALYIELD")),2,-1)%></div></td>
      <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>
      <td><div align="center">
        <%if rs("OVERALL_EXCEPTION")="1" then%>
        <img src="/Images/Yes.gif" width="16" height="16">
        <%else%>
        &nbsp;
        <%end if%>
      </div></td>
      <td><div align="center">
        <%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
        <%=newtime%></div></td>
      <td><div align="center">
        <%unit=unitconvert(csng(rs("WIP_TIME")),newtime)%>
        <%=newtime%></div></td>
      <td><div align="center"><%= rs("CAPACITY")%></div></td>
      <td><div align="center">
        <%if not isnull(rs("INCLUDED_SERIES")) and rs("INCLUDED_SERIES")<>"" then%>
        <% =formatlongstringbreak(getSeries("TEXT",""," where S.NID in ('"&replace(rs("INCLUDED_SERIES"),",","','")&"')"," order by S.SERIES_NAME"," ; "),"<br>",30)%>
        <%else%>
        &nbsp;
        <%end if%>
      </div></td>
      <td><div align="center">
        <%if rs("INCLUDED_SYSTEM_ITEMS")<>"" then%>
        <%=formatlongstringbreak(highlightsamestring(rs("INCLUDED_SYSTEM_ITEMS"),","),"<br>",30)%>
        <%end if%>
        &nbsp;</div></td>
    </tr>
    <%
i=i+1
rs.movenext
wend
%>
    <tr>
      <td height="20" colspan="18"><div align="center">
        <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
        <input name="path" type="hidden" id="path" value="<%=path%>">
        <input name="query" type="hidden" id="query" value="<%=query%>">
        <input name="CheckAll" type="button" id="CheckAll" onClick="checkall()" value="Check All">
        &nbsp;
        <input name="UncheckAll" type="button" id="UncheckAll" onClick="uncheckall()" value="Uncheck All">
        &nbsp;
        <input name="Update" type="submit" id="Update" value="Update">
        &nbsp;
        <input name="Reset" type="reset" id="Reset" value="Reset">
      </div></td>
    </tr>
    <%
else
%>
  </form>
  <tr>
    <td height="20" colspan="17"><div align="center">No Records&nbsp;</div></td>
  </tr>
  <%end if
rs.close%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->