<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
pagename="/System/Task/Task.asp"
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
SQL="select * from TASK order by NID"
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
    <td height="20" colspan="21" class="t-c-greenCopy">Browse Task List </td>
  </tr>
  <tr>
    <td height="20" colspan="21" class="t-c-greenCopy"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="white">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/System/Task/NewTask.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Task</a> </div></td>
      </tr>
    </table>   </td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="21"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center">Index</div></td>
    <td height="20" colspan="2"><div align="center">Action</div></td>
    <td height="20"><div align="center">Name</div></td>
    <td height="20"><div align="center">Chinese Name </div></td>
    <td height="20"><div align="center">Description</div></td>
    <td height="20"><div align="center">Packages</div></td>
    <td><div align="center">Apply</div></td>
    <td height="20"><div align="center">Control Type 1</div></td>
    <td height="20"><div align="center">Parameter 1</div></td>
    <td height="20"><div align="center">Chinese Parameter 1 </div></td>
    <td height="20"><div align="center">Control Type 2</div></td>
    <td height="20"><div align="center">Parameter 2 </div></td>
    <td height="20"><div align="center">Chinese Parameter 2 </div></td>
    <td height="20"><div align="center">Control Type 3 </div></td>
    <td height="20"><div align="center">Parameter 3 </div></td>
    <td height="20"><div align="center">Chinese Parameter 3 </div></td>
    <td height="20"><div align="center">Control Type 4 </div></td>
    <td height="20"><div align="center">Parameter 4 </div></td>
    <td height="20"><div align="center">Chinese Parameter 4 </div></td>
    <td height="20"><div align="center">Engineer</div></td>
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
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('/System/Task/EditTask.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this task? If deleted, task function can not be apply. ')){window.open('/System/Task/DeleteTask.asp?id=<%=rs("NID")%>&taskname=<%=rs("TASK_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" width="16" height="20"></span></div></td>
    <td height="20"><div align="center"><%=rs("TASK_NAME")%></div></td>
    <td height="20"><div align="center"><%=rs("TASK_CHINESE_NAME")%></div></td>
    <td height="20"><div align="center"><%=rs("DESCRIPTION")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PACKAGE")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("APPLY_BY_USER")%></div></td>
    <td height="20"><div align="center"><%=rs("PARAM_TYPE1")%></div></td>
    <td height="20"><div align="center"><%=rs("PARAM1")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_CHINESE1")%>&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;<%=rs("PARAM_TYPE2")%></div></td>
    <td height="20"><div align="center"><%=rs("PARAM2")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_CHINESE2")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_TYPE3")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM3")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_CHINESE3")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_TYPE4")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM4")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("PARAM_CHINESE4")%>&nbsp;</div></td>
    <td height="20"><div align="center"><%'=getRoleMember(true,"TEXT",""," where EVENTS_ID like '%"&rs("NID")&"%'","",""," ; ")%>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="21"><div align="center">No Records </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
