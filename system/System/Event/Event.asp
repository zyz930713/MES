<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
pagename="/System/Event/Event.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
SQL="select * from EVENT order by NID"
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
    <td height="20" colspan="8" class="t-c-greenCopy">Browse Event List </td>
  </tr>
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="white">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Event/NewEvent.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Event</a> </div></td>
      </tr>
    </table>   </td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="8"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center">Index</div></td>
    <td><div align="center">Edit</div></td>
    <td><div align="center">Delete</div></td>
    <td height="20"><div align="center">Name</div></td>
    <td><div align="center">Description</div></td>
    <td><div align="center">Note</div></td>
    <td><div align="center">Body</div></td>
    <td><div align="center">Engineer</div></td>
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
        <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('/System/Event/EditEvent.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this event? If deleted, event can not be apply. ')){window.open('/System/Event/DeleteEvent.asp?id=<%=rs("NID")%>&eventname=<%=rs("EVENT_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" width="16" height="20"></span></div></td>
    <td height="20"><div align="center"><%=rs("EVENT_NAME")%></div></td>
    <td><div align="center"><%=rs("DESCRIPTION")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("EVENT_NOTE")%>&nbsp;</div></td>
    <td><div align="center"><%=formatlongstring(rs("EMAIL_BODY"),"<BR>",60)%>&nbsp;</div></td>
    <td><div align="center"><%=getRoleMember(true,"TEXT",""," where EVENTS_ID like '%"&rs("NID")&"%'","",""," ; ")%>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="8"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
