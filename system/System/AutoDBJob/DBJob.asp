<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
linename=request("linename")
factory=request.QueryString("factory")
if linename<>"" then
where=where&" and lower(L.LINE_NAME) like '%"&lcase(linename)&"%'"
end if

pagename="/System/AutoDBJob/DBJob.asp"
pagepara="&linename="&linename&"&factory="&factory
SQL="select * from user_jobs"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KES CMMS System </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/System/AutoDBJob/Lan_DBJob.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span id="inner_User"></span>:
          <% =session("User") %></td>
        <td><div align="right"><a href="/System/AutoDBJob/AddDBJob.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_Add"></span></a></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_JobID"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_What"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_RunTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_TotalTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Interval"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Failures"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(session("strpagenum")-1)*recordsize+i%>
  </div></td>
    <td height="20"><div align="center">
      <span style="cursor:hand" onClick="javascript:window.open('EditDBJob.asp?id=<%=rs("JOB")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span>
    </div></td>
    <td height="20"><div align="center">
      <span style="cursor:hand" onClick="javascript:window.open('DeleteDBJob.asp?id=<%=rs("JOB")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconDelete.gif" alt="Click to delete"></span>
    </div></td>
    <td class="red"><div align="center">
      <%if rs("BROKEN")="N" then%>
      <span style="cursor:hand" onClick="javascript:location.href='DisableDBJob.asp?id=<%=rs("JOB")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable DB Job"><img src="/Images/Enabled.gif" width="10" height="10"></span>
      <%else%>
      <span style="cursor:hand" onClick="javascript:location.href='EnableDBJob.asp?id=<%=rs("JOB")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable DB Job"><img src="/Images/Disabled.gif"></span>
      <%end if%>
    </div></td>
    <td height="20"><div align="center"><%= rs("JOB") %></div></td>
    <td><div align="center"><%=formatlongstringbreak(rs("WHAT"),"<br>",80)%></div></td>
    <td><div align="center"><%= rs("NEXT_DATE") %></div></td>
    <td><div align="center"><%'= clng(rs("TOTAL_TIME")) %>&nbsp;</div></td>
    <td><div align="center"><%= rs("INTERVAL") %></div></td>
    <td height="20"><div align="center"><%= rs("FAILURES") %>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="10">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->