<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Machine/MachineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Devices/PC/PC.asp"
SQL="select * from BARCODE_DESKTOP order by NID"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy">Browse Desktop List </td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/Machine/AddMachine.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Scanner
        <input type="button" name="Button" value="Import Desktops" onClick="javascript:window.open('ImportPC.asp','','')">
      </a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td width="45" height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <%if admin=true then%>
  <td width="44" height="20" class="t-t-Borrow"><div align="center">Edit</div></td>
  <td width="54" height="20" class="t-t-Borrow"><div align="center">Delete</div></td>
  <%end if%>
  <td width="196" height="20" class="t-t-Borrow"><div align="center">Computer Name </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">Service Tag </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">MAC Address </div></td>
  <td width="196" class="t-t-Borrow"><div align="center">Model</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*pagesize_s+i%>
      </div></td>
    <td height="20"><div align="center" class="red">
      <div align="center">
        <span style="cursor:hand" onClick="javascript:window.open('EditMachine.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">Edit</span>
      </div>
    </div></td>
    <td height="20"><div align="center" class="red">
      <div align="center">
        <span style="cursor:hand" onClick="javascript:window.open('DeleteMachine.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&computername=<%=rs("COMPUTER_NAME")%>','main')">Delete</span>
      </div>
    </div></td>
    <td height="20"><div align="center"><%=rs("COMPUTER_NAME")%></div></td>
    <td><div align="center"><%=rs("SERVICE_TAG")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("MAC_ADDRESS")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("MODEL")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->