<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/GlueWIRE/GlueWIRECheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->


<%
'path=request.ServerVariables("PATH_INFO")
'query=request.ServerVariables("QUERY_STRING")
'query=replace(query,"&","*")
modelname=trim(request("modelname"))
'pagename="/Admin/GlueWIRE/GlueWIRE_Station_List.asp"
where=""
if modelname<>"" then
where=where&" where ITEM_NAME like '%"&modelname&"%'"
end if



pagepara="&modelname="&modelname
sql="select a.station_no,a.ITEM_NAME,b.SUPPLIER_NAME,a.STATION_DESCRIPTION,a.STATION_DESCRIPTION_EN from MATERIAL_STATION a left join Product_model b on( a.ITEM_NAME =b.ITEM_NAME )" &where&" order by STATION_NO "
'SQL="select * from MATERIAL_STATION "&where&" order by STATION_NO "

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
<form method="post" name="form1" target="_self" action="/Admin/GlueWIRE/GlueWIRE_Station_List.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
    <td height="20" colspan="7" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
<tr align="center">
    <td width="4%"  height="20"><span id="inner_SearchPartNumber"></span></td>
    <td width="15%" ><input name="modelname" type="text" id="modelname" value="<%=modelname%>"></td>
	<td width="81%" align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
</tr>	 
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">  
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/GlueWIRE/AddGlueWIRE.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow" ><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StationName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_SubSeries"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
   <td class="t-t-Borrow"><div align="center"><span id="inner_SupplierName"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20" ><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditGlueWIRE.asp?STATION_NO=<%=rs("STATION_NO")%>&path=<%=path%>&query=<%=query%>&Action=Edit','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('EditGlueWIRE1.asp?STATION_NO=<%=rs("STATION_NO")%>&path=<%=path%>&query=<%=query%>&Action=Del','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("STATION_DESCRIPTION") %></div></td>
	<td height="20"><div align="center"><%= rs("STATION_DESCRIPTION_EN")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("ITEM_NAME")%>&nbsp;</div></td>
    <td height="20"><%=rs("SUPPLIER_NAME")%>&nbsp;</td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="7" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if

rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->