<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="/Job/SubJobs/Defect_Change_History.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
code=request("code")
action=request("action")
if action<>"" then
where=where&" and ACTION like '%"&action&"%'"
end if
pagepara="&action="&action
SQL="select * from SYSTEM_LOG where 1=1 "&where&" order by OCCURRED_TIME desc"

 
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<html>
<head>
<title>Defect Code Change List</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy">Browse Log List </td>
  </tr>
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="white">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Role/NewRole.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"></a> </div></td>
      </tr>
    </table>   </td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="6"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center">Index</div></td>
    <td height="20"><div align="center">Occurred Time </div></td>
    <td><div align="center">Module</div></td>
    <td><div align="center">Action</div></td>
    <td><div align="center">Code</div></td>
    <td><div align="center">Engineer</div></td>
  </tr>
  <form name="checkform" method="post" action="/System/Audit/BunchDeleteAudit.asp">
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td height="20"><div align="left"><%=rs("OCCURRED_TIME")%></div></td>
    <td><div align="center"><%=rs("MODULE")%>&nbsp;</div></td>
    <td><%=rs("ACTION")%></td>
    <td><div align="center"><%=rs("USER_CODE")%></div></td>
    <td><div align="center"><%=rs("USER_NAME")%>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend%>
</form>
<%else%>
  <tr>
    <td height="20" colspan="6"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->