<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
linename=request("linename")
result=request("result")
factory=request.QueryString("factory")
if linename<>"" then
where=where&" and lower(L.LINE_NAME) like '%"&lcase(linename)&"%'"
end if
if result<>"" and result<>"All" then
where=where&" and L.SUCCESS='"&result&"'"
end if
pagename="/Job/Shift/ScheduleRecords.asp"
pagepara="&linename="&linename&"&result="&result
if session("code")<>"1194" then
SQL="select 1,L.*,U.USER_CHINESE_NAME from LINE_SHIFT_SCHEDULE L left join USERS U on L.PERSON=U.USER_CODE where L.PERSON='"&session("code")&"' "&where&factorywhereoutsideand&" order by L.LINE_NAME,L.RUNTIME desc"
else
SQL="select 1,L.*,U.USER_CHINESE_NAME from LINE_SHIFT_SCHEDULE L left join USERS U on L.PERSON=U.USER_CODE where 1=1 "&where&factorywhereoutsideand&" order by L.LINE_NAME,L.RUNTIME desc"
end if
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/Shift/Lan_ScheduleRecords.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Job/Shift/ScheduleRecords.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchLine"></span></td>
    <td><input name="linename" type="text" id="linename" value="<%=linename%>"></td>
    <td><span id="inner_SearchResult"></span></td>
    <td><select name="result" id="result">
      <option>All</option>
      <option value="1" <%if result="1" then%>selected<%end if%>>Success</option>
      <option value="2" <%if result="2" then%>selected<%end if%>>Fail</option>
    </select>
    </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_LineName"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Person"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ShiftType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Runtime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_RunResult"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ErrorInfo"></span></div></td>
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
    <td><div align="center"><%if cdate(rs("RUNTIME"))<=now() then%><img src="/Images/IconEdit_No.gif"><%else%><span class="red" style="cursor:hand" onClick="javascript:window.open('EditSchedule.asp?job_id=<%=rs("JOB_ID")%>&line_name=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span><%end if%></div></td>
    <td><%if cdate(rs("RUNTIME"))<=now() then%><img src="/Images/IconDelete_No.gif" width="16" height="20"><%else%><span class="red" style="cursor:hand" onClick="javascript:window.open('DeleteSchedule.asp?job_id=<%=rs("JOB_ID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconDelete.gif" width="16" height="20"></span><%end if%></td>
    <td><div align="center">
      <%if rs("SUCCESS")="1" then%>成功<%elseif rs("SUCCESS")="2" then%>失败<%else%>未开始<%end if%>
    </div></td>
    <td height="20"><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td><div align="center"><%= rs("USER_CHINESE_NAME")%></div></td>
    <td><div align="center"><%if rs("SHIFT_TYPE")="STOP" then%>停线<%else%>开线<%end if%></div></td>
    <td><div align="center"><%= rs("RUNTIME")%></div></td>
    <td height="20"><div align="center"><%= rs("RUN_RESULT")%>&nbsp;</div></td>
    <td><%= rs("ERROR_INFO")%>&nbsp;</td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="10"><div align="center">No Records&nbsp;</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->