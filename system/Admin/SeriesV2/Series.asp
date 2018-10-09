<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
seriesname=request("seriesname")
thisstatus=request("status")
factory=request("factory")
if seriesname<>"" then
where=where&" and lower(SR.SERIES_NAME) like '%"&lcase(seriesname)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and SR.STATUS="&thisstatus
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and SR.FACTORY_ID is null"
else
where=where&" and SR.FACTORY_ID='"&factory&"'"
end if

pagename="/Admin/Series/Series.asp"
pagepara="&seriesname="&seriesname&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "SR."
SQL="select 1,SR.*,F.FACTORY_NAME,S.SECTION_NAME,L.LINE_NAME from SERIES SR inner join FACTORY F on SR.FACTORY_ID=F.NID left join SECTION S on SR.SECTION_ID=S.NID left join LINE L on SR.LINE_ID=L.NID where 1=1 "&where&factorywhereoutsideand&" order by SR.SERIES_NAME"
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
<form action="/Admin/Series/Series.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn">Search Series</td>
  </tr>
  <tr>
    <td width="11%" height="20">Series Name </td>
    <td width="18%"><input name="seriesname" type="text" id="seriesname" value="<%=seriesname%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">Browse Series List </td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Series_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Series/AddSeries.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Series </a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="3" class="t-t-Borrow"><div align="center">Action</div></td>
  <td class="t-t-Borrow"><select name="status" id="status" onChange="location.href='<%=pagename%>?status='+this.options[this.selectedIndex].value+'&station=<%=station%>&factory=<%=factory%>&purpose=<%=purpose%>'">
    <option value="">Status</option>
    <option value="all" <%if thisstatus="all" then%>selected<%end if%>>All</option>
    <option value="1" <%if thisstatus="1" then%>selected<%end if%>>Enabled</option>
    <option value="0" <%if thisstatus="0" then%>selected<%end if%>>Disabled</option>
  </select></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Series Name </div></td>
  <td class="t-t-Borrow"><div align="center">Line</div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">Section</div></td>
  <td class="t-t-Borrow"><div align="center">Internal Yield </div></td>
  <td class="t-t-Borrow"><div align="center">Target Yield </div></td>
  <td class="t-t-Borrow"><div align="center">Exception</div></td>
  <td class="t-t-Borrow"><div align="center">Included models </div></td>
  </tr>
<%
all_model=""
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*pagesize_s+i%>
  </div></td>
  <%if admin=true then%>
  <td><div align="center"><span style="cursor:hand" onClick="javascript:window.open('CopySeries.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
    <img src="/Images/IconCopy.gif" alt="Click to copy"></span>&nbsp;</div></td>
    <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditSeries.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('DeleteSeries.asp?id=<%=rs("NID")%>&series_name=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
	<td><div align="center"><span class="red">
	  <%if rs("STATUS")="1" then%>
	  <span style="cursor:hand" onClick="javascript:location.href='DisableSeries.asp?id=<%=rs("NID")%>&seriesname=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Series"><img src="/Images/Enabled.gif"></span>
	  <%else%>
	  <span style="cursor:hand" onClick="javascript:location.href='EnableSeries.asp?id=<%=rs("NID")%>&seriesname=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Series"><img src="/Images/Disabled.gif"></span>
	  <%end if%>
    </span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("SERIES_NAME") %></div></td>
    <td><div align="center"><%= rs("LINE_NAME") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
    <td><div align="center"><%= formatpercent(csng(rs("TARGET_INTERNALYIELD")),2,-1)%></div></td>
    <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>
    <td><div align="center"><%if rs("OVERALL_EXCEPTION")="1" then%><img src="/Images/Yes.gif" width="16" height="16"><%else%>&nbsp;<%end if%></div></td>
    <td><div align="center"><%if rs("INCLUDED_SYSTEM_ITEMS")<>"" then%><%=highlightsamestring(rs("INCLUDED_SYSTEM_ITEMS"),",")%><%end if%>&nbsp;</div></td>
  </tr>
<%
all_model=all_model&rs("INCLUDED_SYSTEM_ITEMS")&","
i=i+1
rs.movenext
wend
all_model=left(all_model,len(all_model)-1)
%>
  <tr>
    <td height="20" colspan="13"><%=showsamestring(all_model,",")%>&nbsp;</td>
  </tr>
<%
else
%>
  <tr>
    <td height="20" colspan="13">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->