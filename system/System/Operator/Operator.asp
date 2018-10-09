<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
pagename="/System/Operator/Operator.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
SQL="select * from OPERATORS order by CODE"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy">Browse Operator List </td>
  </tr>
  <tr class="t-c-greenCopy">
    <td height="20" colspan="6"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="white">User:
            <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Operator/NewOperator.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Operator</a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr bordercolorlight="#000099">
    <td class="t-t-Borrow"><div align="center">Index</div></td>
    <td class="t-t-Borrow"><div align="center">Edit</div></td>
    <td class="t-t-Borrow"><div align="center">Delete</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">User Code</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">User Name</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">Authorized Stations </div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
<tr bordercolorlight="#000099">
    <td><div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditOperator.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">Edit</span></div></td>
    <td><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('DeleteOperator.asp?id=<%=rs("NID")%>&operatorname=<%=rs("OPERATOR_NAME")%>&path=<%=path%>&query=<%=query%>','main')">Delete</span></div></td>
    <td height="20"><div align="center"><%= rs("CODE") %></div></td>
    <td height="20"><div align="center"><%= rs("OPERATOR_NAME") %></div></td>
    <td><div align="left">
        <%if rs("AUTHORIZED_STATIONS_ID")<>"" then%>
        <%= getStation(true,"TEXT",""," where S.NID in ('"&replace(rs("AUTHORIZED_STATIONS_ID"),",","','")&"')","",rs("AUTHORIZED_STATIONS_ID")," ; ") %>
        <%end if%>
&nbsp;</div></td>
  </tr>

  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="6"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
