<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
JOB_TYPE=request("txtJOB_TYPE")
JOB_TYPE_DESC=request("txtJOB_TYPE_DESC")


if(request.QueryString("JOB_TYPE")<>"") then
	set rsPRIORITY=server.createobject("adodb.recordset")
	JOB_TYPE=request.QueryString("JOB_TYPE")
	SQL="DELETE JOB_TYPE_SETTING WHERE JOB_TYPE='"+JOB_TYPE+"'"
	rsPRIORITY.open SQL,conn,1,3
end if 

where=" where 1=1"
if(JOB_TYPE<>"") then
	where=where+" and JOB_TYPE like '"+JOB_TYPE+"%'" 
end if 

if(JOB_TYPE_DESC<>"") then
	where=where+" and JOB_TYPE_DESC like '"+JOB_TYPE_DESC+"%'" 
end if 

SQL="select * from JOB_TYPE_SETTING "&where&"  order by  JOB_TYPE"

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
<form action="/Admin/JobType/JobTypeList.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search Job Type</td>
  </tr>
  <tr>
    <td height="20">Job Type</td>
    <td><input name="txtJOB_TYPE" type="text" id="txtJOB_TYPE" value="<%=JOB_TYPE%>"></td>
    <td height="20">Job Type Description</td>
    <td><input name="txtJOB_TYPE_DESC" type="text" id="txtJOB_TYPE_DESC" value="<%=JOB_TYPE_DESC%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy">Browse Job Type</td>
</tr>
<tr>
  <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/JobType/JobType.asp" target="main" class="white">Add a Job Type</a><%end if%></div></td>
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
  <td class="t-t-Borrow"><div align="center">Job Type</div></td>
  <td class="t-t-Borrow"><div align="center">Job Type Description</div></td>
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
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('JobType.asp?Action=edit&JOB_TYPE=<%=rs("JOB_TYPE")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Job Type?')){form1.action='JobTypeList.asp?JOB_TYPE=<%=rs("JOB_TYPE")%>';form1.submit();}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
	<%end if%>
    <td><div align="center"><%= rs("JOB_TYPE") %></div></td>
    <td><div align="center"><%= rs("JOB_TYPE_DESC") %></div></td>
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