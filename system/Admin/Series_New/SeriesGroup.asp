<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
seriesgroupname=request("seriesgroupname")
thisstatus=request("status")
factory=request("factory")
family=request("family")
if seriesgroupname<>"" then
where=where&" and lower(SG.SERIES_NAME) like '%"&lcase(seriesgroupname)&"%'"
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
if not isnull(family) and family<>"" then
where=where&" and FA.NID='"&family&"'"
end if


pagename="/Admin/SERIES_NEW/SeriesGroup.asp"
pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "SG."
SQL="select SG.*,F.FACTORY_NAME,S.SECTION_NAME,FA.FAMILY_NAME from SERIES_NEW SG inner join FACTORY F on SG.FACTORY_ID=F.NID left join SECTION S on SG.SECTION_ID=S.NID left join Family FA on SG.FAMILY_ID=FA.NID where 1=1 "&where&factorywhereoutsideand&" order by SG.SERIES_NAME"
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
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Admin/SERIES_NEW/SeriesGroup.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr align="center">
    <td height="20" width="80"><span id="inner_Series"></span></td>
    <td width="100"><input name="seriesgroupname" type="text" id="seriesgroupname" value="<%=seriesgroupname%>"></td>
    <td width="60"><span id="inner_SearchStatus"></span></td>
    <td width="60"><span class="t-t-Borrow">
      <select name="status" id="status">
        <option value=""></option>
        <option value="1" <%if thisstatus="1" then%>selected<%end if%>>Enabled</option>
        <option value="0" <%if thisstatus="0" then%>selected<%end if%>>Disabled</option>
      </select>
    </span></td>
    
  	<Td width="60"><span id="inner_SearchFamily"></span></Td>
  	 <td width="80"><select name="family" id="family">
    <option value=""></option>
    <%= getFamily("OPTION",family,factorywhereoutside,"","") %>
  </select>  </td>
  	<Td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></Td>

  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%"><span id="inner_BrowseData"></span></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('SeriesGroup_Export.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
        <td width="50%"><div align="right">
          <%if admin=true then%>
          <a href="/Admin/SERIES_NEW/AddSeriesGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
          <%end if%>
        </div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="18"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <%if admin=true then%>
<!--	
    <td height="20" class="t-t-Borrow"><div align="center">Select</div></td>
-->
    <td height="20" colspan="3" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
    <%end if%>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Series"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_Factory"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_Section"></span></div></td>
	<td class="t-t-Borrow"><div align="center"><span id="td_FirstPassYield"></span></td>	
    <td class="t-t-Borrow"><div align="center"><span id="td_InternalYield"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_TargetYield"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime"></span> </div></td>
	<!--
    <td class="t-t-Borrow"><div align="center">WIP Time </div></td>
    <td class="t-t-Borrow"><div align="center">Capacity</div></td>
	-->
	<td class="t-t-Borrow"><div align="center"><span id="td_Family"></span></div></td>
	<td class="t-t-Borrow"><div align="center"><span id="td_SubSeries"></span></div></td>
  </tr>
  <form name="checkform" method="post" action="/Admin/SERIES_NEW/UpdateSeriesGroupAll.asp">
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
<!--	  
      <td><div align="center">
        <input name="nid<%=i%>" type="hidden" id="nid<%=i%>" value="<%=rs("NID")%>">
        <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">
-->
      </div></td>
      <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditSeriesGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
      <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Series Group?')){window.open('DeleteSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
      <td><div align="center"><span class="red">
        <%if rs("STATUS")="1" then%>
        <span style="cursor:hand" onClick="javascript:location.href='DisableSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Series Group"><img src="/Images/Enabled.gif"></span>
        <%else%>
        <span style="cursor:hand" onClick="javascript:location.href='EnableSeriesGroup.asp?id=<%=rs("NID")%>&seriesgroupname=<%=rs("SERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Series Group"><img src="/Images/Disabled.gif"></span>
        <%end if%>
      </span></div></td>
      <%end if%>
      <td height="20"><div align="center"><%= rs("SERIES_NAME") %></div></td>
      <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
      <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
	  <td><div align="center"><%= csng(rs("TARGET_FIRSTYIELD"))%>%</div></td>
      <td><div align="center"><%= formatpercent(csng(rs("TARGET_INTERNALYIELD")),2,-1)%></div></td>
      <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>      
      <td><div align="center">
        <%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
        <%=newtime%></div></td>
	<!--	
      <td><div align="center">
        <%unit=unitconvert(csng(rs("WIP_TIME")),newtime)%>
        <%=newtime%></div></td>
      <td><div align="center"><%= rs("CAPACITY")%></div></td>
	  -->
	  <td><div align="center"><%= rs("FAMILY_NAME")%></div></td>
	    <td><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('ShowIncludeSubSeries.asp?id=<%=rs("NID")%>')"><img src="/Images/IconEdit.gif" ></span></div></td>
      
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
		<!--
        <input name="CheckAll" type="button" id="CheckAll" onClick="checkall()" value="Check All">
        &nbsp;
        <input name="UncheckAll" type="button" id="UncheckAll" onClick="uncheckall()" value="Uncheck All">
        &nbsp;
        <input name="Update" type="submit" id="Update" value="Update">
        &nbsp;
        <input name="Reset" type="reset" id="Reset" value="Reset">
		-->
      </td>
    </tr>
    <%
else
%>
  </form>
  <tr>
    <td height="20" colspan="17"><div align="center"><span id="inner_Records"></span>&nbsp;</div></td>
  </tr>
  <%end if
rs.close%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->