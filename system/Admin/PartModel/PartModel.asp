<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
model_number=request.QueryString("model_number")
if model_number<>"" then
where=where&" and lower(MODEL_NUMBER) like '%"&lcase(model_number)&"%'"
end if

pagename="/Admin/Line/Line.asp"
pagepara="&model_number="&model_number&"&factory="&factory
SQL="select * from PART_MODEL where ITEM_ID is not null "&where&" order by MODEL_NUMBER"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Admin/PartModel/Lan_PartModel.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Admin/PartModel/PartModel.asp" method="get" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchModelNumber"></span></td>
    <td><input name="model_number" type="text" id="model_number" value="<%=model_number%>"></td>
    <td><span id="inner_SearchFactory"></span></td>
    <td><select name="factory" id="factory">
      <option value="">Factory</option>
    </select></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td width="54" height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td width="241" height="20" class="t-t-Borrow"><div align="center"><span id="inner_ModelNumber"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_Description"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_WIPInventory"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_LeadTime"></span></div></td>
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
    <td width="51" height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditPartModel.asp?id=<%=rs("ITEM_ID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <%end if%>
    <td height="20"><div align="center"><%= rs("MODEL_NUMBER") %></div></td>
    <td><div align="center"><%= rs("DESCRIPTION")%></div></td>
    <td><div align="center"><%= rs("WIP_SUPPLY_SUBINVENTORY")%>&nbsp;</div></td>
    <td><div align="center"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%><%=newtime&" "&unit%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="6">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->