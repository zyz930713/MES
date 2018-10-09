<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")

pagename="/Job/Shift/CorrectShiftTime.asp"
SQL="select SHIFT_IN_TIME,SHIFT_OUT_TIME,SHIFT_IN_PERSON,SHIFT_OUT_PERSON from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
%>
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
  <td height="20" colspan="3" class="t-c-greenCopy">工单号：<%=jobnumber&"-"&repeatstring(sheetnumber,"0",3)%></td>
</tr>
<tr>
  <td height="20" colspan="3" class="t-c-greenCopy">用户:<% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="3">&nbsp;</td>
</tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center">序号</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">开线时间</div></td>
  <td class="t-t-Borrow"><div align="center">停线时间</div></td>
  </tr>
<%
if not rs.eof then
a_shift_in_time=split(rs("SHIFT_IN_TIME"),",")
a_shift_in_person=split(rs("SHIFT_IN_PERSON"),",")
a_shift_out_time=split(rs("SHIFT_OUT_TIME"),",")
a_shift_out_person=split(rs("SHIFT_OUT_PERSON"),",")
if ubound(a_shift_in_time)>=ubound(a_shift_out_time) then
lastindex=ubound(a_shift_in_time)
else
lastindex=ubound(a_shift_out_time)
end if
	for j=0 to lastindex-1
%>
<tr>
  <td height="20"><div align="center">
    <% =j+1%>
  </div></td>
    <td height="20"><div align="center"><%if j<ubound(a_shift_in_time) then%><span style="cursor:hand" title="" onClick="javascript:window.open('CorrectShiftTime1.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>&sequence=<%=j%>&shifttype=in&path=<%=path%>&query=<%=query%>','_self')"><img src="/Images/IconDelete.gif" width="16" height="20">&nbsp;<%= a_shift_in_time(j) %>&nbsp;(<%=a_shift_in_person(j)%>)</span><%else%>&nbsp;<%end if%></div></td>
    <td><div align="center"><%if j<ubound(a_shift_out_time) then%><span style="cursor:hand" title="" onClick="javascript:window.open('CorrectShiftTime1.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>&sequence=<%=j%>&shifttype=out&path=<%=path%>&query=<%=query%>','_self')"><img src="/Images/IconDelete.gif" width="16" height="20">&nbsp;<%= a_shift_out_time(j) %>&nbsp;(<%=a_shift_out_person(j)%>)</span><%else%>&nbsp;<%end if%></div></td>
  </tr>
<%
	next
else
%>
  <tr>
    <td height="20" colspan="3">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->