<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
PriorityLevel=request("txtPriorityLevel")
PriorityDescription=request("txtPriorityDescription")
PRIORITY_LEVEL=request.QueryString("PRIORITY_LEVEL")

if(PRIORITY_LEVEL<>"") then
	set rsPRIORITY=server.createobject("adodb.recordset")
	SQL="DELETE JOB_PRIORITY_SETTING WHERE PRIORITY_LEVEL='"+PRIORITY_LEVEL+"'"
	rsPRIORITY.open SQL,conn,1,3
end if 

where=" where 1=1"
if(PriorityLevel<>"") then
	where=where+" and PRIORITY_LEVEL like '"+PriorityLevel+"%'" 
end if 

if(PriorityDescription<>"") then
	where=where+" and PRIORITY_DEC like '"+PriorityDescription+"%'" 
end if 

SQL="select * from JOB_PRIORITY_SETTING "&where&"  order by  PRIORITY_LEVEL"

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
<form action="/Admin/JobPrioritySetting/JobPriorityList.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search Job Priority</td>
  </tr>
  <tr>
    <td height="20">Priority Level</td>
    <td><input name="txtPriorityLevel" type="text" id="txtPriorityLevel" value="<%=PriorityLevel%>"></td>
    <td height="20">Priority Description</td>
    <td><input name="txtPriorityDescription" type="text" id="txtPriorityDescription" value="<%=PriorityDescription%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">Browse Priority</td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/JobPrioritySetting/JobPriority.asp" target="main" class="white">Add a New Priority</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center">Priority Level</div></td>
  <td class="t-t-Borrow"><div align="center">Priority Description</div></td>
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
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('JobPriority.asp?Action=edit&PRIORITY_LEVEL=<%=rs("PRIORITY_LEVEL")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Job Priority Level?')){form1.action='JobPriorityList.asp?PRIORITY_LEVEL=<%=rs("PRIORITY_LEVEL")%>';form1.submit();}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	<%end if%>
    <td><div align="center"><%= rs("PRIORITY_LEVEL") %></div></td>
    <td><div align="center"><%= rs("PRIORITY_DEC") %></div></td>
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