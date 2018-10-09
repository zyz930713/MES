<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/FinanceSeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
seriesgroupname=request("seriesgroupname")
thisstatus=request("status")
if seriesgroupname<>"" then
where=where&" and lower(SG.SERIES_GROUP_NAME) like '%"&lcase(seriesgroupname)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and SG.STATUS="&thisstatus
end if

pagename="/Admin/FinanceSeriesGroup/SeriesGroup.asp"
pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
FactoryRight "SG."
SQL="select SG.* from FINANCE_SERIES_GROUP SG where SG.NID is not null "&where&factorywhereoutsideand&" order by SG.SERIES_GROUP_NAME"
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
    <td height="20" colspan="3" class="t-b-midautumn">Search Finance Series Group</td>
  </tr>
  <tr>
    <td width="13%" height="20">Finance Series Group Name </td>
    <td width="16%"><input name="seriesgroupname" type="text" id="seriesgroupname" value="<%=seriesgroupname%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">Browse Finance Series Group List </td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('SeriesGroup_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/FinanceSeriesGroup/AddSeriesGroup.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Finance Series Group</a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="status" id="status" onChange="location.href='<%=pagename%>?status='+this.options[this.selectedIndex].value+'&station=<%=station%>&factory=<%=factory%>&purpose=<%=purpose%>'">
      <option value="">Status</option>
      <option value="all" <%if thisstatus="all" then%>selected<%end if%>>All</option>
      <option value="1" <%if thisstatus="1" then%>selected<%end if%>>Enabled</option>
      <option value="0" <%if thisstatus="0" then%>selected<%end if%>>Disabled</option>
    </select>
  </div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Series Group Name </div></td>
  <td class="t-t-Borrow"><div align="center">Target Yield </div></td>
  <td class="t-t-Borrow"><div align="center">Plan Target </div></td>
  <td class="t-t-Borrow"><div align="center">Prefix</div></td>
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
  <%if admin=true then%>
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
    <td><div align="center"><%= formatpercent(csng(rs("TARGET_YIELD")),2,-1) %></div></td>
    <td><div align="center"><%= formatpercent(csng(rs("PLAN_TARGET_YIELD")),2,-1) %></div></td>
    <td><div align="center"><%=formatlongstringbreak(rs("PREFIX"),"<br>",50)%></div></td>
    <td><div align="center"><%=formatlongstringbreak(getFinanceSeries("TEXT",nul," and SERIES_GROUP_ID='"&rs("NID")&"'"," order by SERIES_NAME"," ; "),"<br>",50)%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
</form> 
<%
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->