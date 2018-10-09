<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="/System/Audit/Audit.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
code=request("code")
action=request("action")
if code<>"" then
where=where&" and USER_CODE='"&code&"'"
end if
if action<>"" then
where=where&" and ACTION like '%"&action&"%'"
end if
pagepara="&code="&code&"&action="&action
SQL="select * from SYSTEM_LOG where 1=1 "&where&" order by OCCURRED_TIME desc"
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
<form name="form1" method="post" action="/System/Audit/Audit.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">Search Engineer</td>
  </tr>
  <tr>
    <td height="20">Code</td>
    <td height="20"><input name="code" type="text" id="code" value="<%=code%>"></td>
    <td>Action</td>
    <td><input name="action" type="text" id="action" value="<%=action%>"></td>
    <td>NT Account </td>
    <td><input name="account" type="text" id="account" value="<% =account%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy">Browse Log List </td>
  </tr>
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="white">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Role/NewRole.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Role</a> </div></td>
      </tr>
    </table>   </td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="8"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center">Index</div></td>
    <td><div align="center">Select</div></td>
    <td><div align="center">Delete</div></td>
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
    <td>
      <div align="center">
        <input name="nid<%=i%>" type="hidden" id="nid<%=i%>" value="<%=rs("NID")%>">
        <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">    
      </div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this log? ')){window.open('/System/Audit/DeleteAudit.asp?id=<%=rs("NID")%>&action=<%=replace(rs("ACTION"),"'","$")%>&path=<%=path%>&query=<%=query%>','main')}">Delete</span></div></td>
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
  <tr>
    <td height="20" colspan="8"><div align="center">
      <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">

      <input name="CheckAll" type="button" id="CheckAll" onClick="checkall()" value="Check All">
      &nbsp;
      <input name="UncheckAll" type="button" id="UncheckAll" onClick="uncheckall()" value="Uncheck All">
      &nbsp;
      <input name="Delete" type="submit" id="Delete" value="Delete">
    &nbsp;  
    <input name="Reset" type="reset" id="Reset" value="Reset">
    </div></td>
  </tr>
</form>
<%else%>
  <tr>
    <td height="20" colspan="8"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->