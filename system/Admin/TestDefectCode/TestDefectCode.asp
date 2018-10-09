<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TestDefectCode/TestDefectCodeCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetMaterial.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Admin/TestDefectCode/TestDefectCode.asp"
station=request.QueryString("station")
where=""
SQL="select * from TEST_DEFECTCODE where 1=1 "&where
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
  <td height="20" colspan="6" class="t-c-greenCopy">Browse Test Defect Code List </td>
</tr>
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/TestDefectCode/AddTestDefectCode.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Test Defect Code</a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Defect Name </div></td>
  <td class="t-t-Borrow"><div align="center">Value Type </div></td>
  <td class="t-t-Borrow"><div align="center">Scale</div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td width="91" height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td width="85" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:window.open ('EditTestDefectCode.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">
<img src="/Images/IconEdit.gif" width="16" height="20"></span></div></td>
    <td width="112" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this Action? If deleted, new job can not apply it forever. ')){window.open('DeleteTestDefectCode.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>&defectcodename=<%=rs("DEFECT_NAME")%>','main')}"><img src="/Images/IconDelete.gif" width="16" height="20"></span></div></td>
	<%end if%>
    <td width="394" height="20"><div align="center"><%= rs("DEFECT_NAME")%></div></td>
    <td width="394"><div align="center"><%= rs("VALUE_TYPE")%></div></td>
    <td width="394"><div align="center"><%= rs("SCALE")%></div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="6"><div align="center">No records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
