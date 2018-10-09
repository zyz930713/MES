<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Factory/FactoryCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<%

if(request.querystring("terminal_part")<>"") then
	TERMINAL_PART=request.querystring("terminal_part")
	set rs0=server.createobject("adodb.recordset")
	SQL="DELETE FROM SOLDER_PROGRAM WHERE  upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"'"
	rs0.open SQL,conn,1,3

	set rs2=server.createobject("adodb.recordset")
	SQL="DELETE FROM SOLDER_PART WHERE  upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"'"
	rs2.open SQL,conn,1,3
	response.write "<script>window.alert('Delete Successfully!');</script>"
end if 

Model_Name=request("txtModelName")
TERMINAL_PART=request("txtTerminalPart")
SQL="select * from SOLDER_PROGRAM WHERE 1=1 "
if(Model_Name <>"") then
	SQL=SQL+" AND  TERMINAL_PART IN (SELECT TERMINAL_PART FROM SOLDER_PART WHERE FG_PART='"+Model_Name+"')"
end if 

if(TERMINAL_PART <>"") then
	SQL=SQL+" AND  TERMINAL_PART='"+TERMINAL_PART+"'"
end if 
SQL=SQL+" ORDER BY TERMINAL_PART"
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
<FORM ID="form1"  name="form1" action="solderlist.asp" method="post">
<body onLoad="language_page()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn">Search</td>
  </tr>
  <tr>
    <td width="11%" height="20">Model Name</td>
	<td width="11%" height="20"><input name="txtModelName" id="txtModelName" type="text"  value="<%=Model_Name%>"></td>
    <td width="18%">Termial Part</td>
	<td width="11%" height="20"><input name="txtTerminalPart" id="txtTerminalPart" type="text"  value="<%=TERMINAL_PART%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy">Browse Solder Program List </td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="SolderProgram.asp" target="main" class="white">Add a New Program </a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="11"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">Edit</div></td>
   <td height="20" class="t-t-Borrow"><div align="center">Delete</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Terminal Part</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Program Number</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('SolderProgram.asp?TERMINAL_PART=<%=rs("TERMINAL_PART")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
   <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Setting?')){form1.action='solderlist.asp?terminal_part=<%=rs("TERMINAL_PART")%>';form1.submit();}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
    <td height="20"><div align="center"><%= rs("TERMINAL_PART") %></div></td>
    <td height="20"><div align="center"><%= rs("PROGRAM_NUMBER") %>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="11">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->