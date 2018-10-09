<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
admin=true
PART_NUMBER=request("txtPartNumber")
where=" where 1=1 "
if PART_NUMBER<>"" then
	where=where&" and PART_NUMBER like '%"&PART_NUMBER&"%'"
end if
if(request.QueryString("Action")="delete") then
	PART_NUMBER=request.QueryString("PART_NUMBER")
	set rsBOMList=server.createobject("adodb.recordset")
	SQL="delete from ITEM_BOM where PART_NUMBER='"+PART_NUMBER+"'"	
	rsBOMList.open SQL,conn,1,3
	where=" where 1=1"
end if 
SQL="select  DISTINCT PART_NUMBER FROM ITEM_BOM " & where
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
<form action="/Admin/SubLineBOM/SubLineBOMList.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-b-midautumn">Search Part Number BOM</td>
  </tr>
  <tr>
  	<td height="20">Part Number</td>
    <td><input name="txtPartNumber" type="text" id="txtPartNumber" value="<%=PART_NUMBER%>"></td>
    <td colspan="5">
	<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">Browse BOM list </td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/SubLineBOM/SubLineBOM.asp" target="main" class="white">Add a New BOM</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow" width="20%"><div align="center">Action</div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center">Part Number</div></td>
   <td class="t-t-Borrow"><div align="center">Detail</div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <%if admin=true then%>
    <td height="20">
		<div align="center" class="red">
			<span style="cursor:hand" onClick="javascript:window.open('SubLineBOM.asp?PART_NUMBER=<%=rs("PART_NUMBER")%>&Action1=edit','main')">
			<img src="/Images/IconEdit.gif" alt="Click to edit"></span></div>
	</td>
	  <td height="20"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this setting?  ')){window.open('SubLineBOMList.asp?action=delete&PART_NUMBER=<%=rs("PART_NUMBER")%>','main')}
	 "><img src="/Images/IconDelete.gif" alt="Click to delete"></span>
	</td>

	<%end if%>
	<td><div align="center"><%= rs("PART_NUMBER") %></div></td>
	 <td height="20">
		<div align="center" class="red">
			<span style="cursor:hand" onClick="javascript:window.open('../../job/SubLinejob/Material_BOM.asp?Material_Part_Number=<%=rs("PART_NUMBER")%>&jobnumber=0')">
			<img src="/Images/IconEdit.gif" alt="Click to edit"></span></div>
	</td>
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