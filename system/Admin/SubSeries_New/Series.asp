<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
seriesname=request("seriesname")
thisstatus=request("status")
factory=request("factory")

Series=request("Series")

if seriesname<>"" then
	where=where&" and lower(SR.SUBSERIES_NAME) like '%"&lcase(seriesname)&"%'"
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

if not isnull(Series) and Series<>""then
	where=where&" and SR.SERIES_ID='"&Series&"'"
end if

pagename="/Admin/SubSeries_New/Series.asp"
pagepara="&seriesname="&seriesname&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "SR."
SQL="select SR.*,F.FACTORY_NAME,S.SECTION_NAME,L.LINE_NAME ,SN.SERIES_NAME from SUBSERIES SR inner join FACTORY F on SR.FACTORY_ID=F.NID left join SECTION S on SR.SECTION_ID=S.NID left join LINE L on SR.LINE_ID=L.NID left join SERIES_NEW SN on SR.SERIES_ID=SN.NID where 1=1 "&where&factorywhereoutsideand&" order by SR.SUBSERIES_NAME"
'response.Write(SQL)
'response.End()
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
<form action="/Admin/SubSeries_New/Series.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_SubSeries"></span></td>
    <td ><input name="seriesname" type="text" id="seriesname" value="<%=seriesname%>"></td>
    <td height="20"><span id="inner_Series"></span></td>
	<td height="20"><select name="Series" id="Series">
    <option value=""></option>
	<%FactoryRight "S."%>
    <%= getSeries("OPTION",Series,factorywhereoutside,"","") %>
  </select>  </td>  
    <td ><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span id="inner_BrowseData"></span></td>        
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/SubSeries_New/AddSeries.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="14"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SubSeries"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Factory"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Section"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_FirstPassYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_InternalYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_TargetYield"></span></div></td>
  <!--
  <td class="t-t-Borrow"><div align="center">Exception</div></td>
  -->
  <td class="t-t-Borrow"><div align="center"><span id="td_Series"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
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
    <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditSeries.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('DeleteSeries.asp?id=<%=rs("NID")%>&series_name=<%=rs("SUBSERIES_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
	<td><div align="center"><span class="red">
	  <%if rs("STATUS")="1" then%>
	  <span style="cursor:hand" onClick="javascript:location.href='DisableSeries.asp?id=<%=rs("NID")%>&seriesname=<%=rs("SUBSERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Series"><img src="/Images/Enabled.gif"></span>
	  <%else%>
	  <span style="cursor:hand" onClick="javascript:location.href='EnableSeries.asp?id=<%=rs("NID")%>&seriesname=<%=rs("SUBSERIES_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Series"><img src="/Images/Disabled.gif"></span>
	  <%end if%>
    </span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("SUBSERIES_NAME") %></div></td>
    <td><div align="center"><%= rs("LINE_NAME") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
	<td><div align="center"><%= rs("TARGET_FIRSTYIELD")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("TARGET_INTERNALYIELD")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>
	<!--
    <td><div align="center"><%if rs("OVERALL_EXCEPTION")="1" then%><img src="/Images/Yes.gif" width="16" height="16"><%else%>&nbsp;<%end if%></div></td>
	-->
	<Td><div align="center"><%= rs("SERIES_NAME") %>&nbsp;</div></Td>
     <td><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.showModalDialog('ShowIncludeModel.asp?id=<%=rs("NID")%>','dialogHeight:600px;dialogWidth:600px')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
  <tr>
    <td height="20" colspan="14">&nbsp;</td>
  </tr>
<%
else
%>
  <tr>
    <td height="20" colspan="14"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->