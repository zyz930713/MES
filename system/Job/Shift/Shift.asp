<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
linename=request("linename")
factory=request.QueryString("factory")
if linename<>"" then
where=where&" and lower(L.LINE_NAME) like '%"&lcase(linename)&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and L.FACTORY_ID is null"
else
where=where&" and L.FACTORY_ID='"&factory&"'"
end if

pagename="/Job/Shift/Shift.asp"
pagepara="&linename="&linename&"&factory="&factory
FactoryRight "L."
if session("code")<>"1194" and session("code")<>"1488" then
SQL="select 1,L.*,F.FACTORY_NAME,S.SECTION_NAME,U.USER_CHINESE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID left join SECTION S on L.SECTION_ID=S.NID left join USERS U on L.LEADER=U.USER_CODE where L.LEADER='"&session("code")&"' "&where&factorywhereoutsideand&" order by L.LINE_NAME"
else
SQL="select 1,L.*,F.FACTORY_NAME,S.SECTION_NAME,U.USER_CHINESE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID left join SECTION S on L.SECTION_ID=S.NID left join USERS U on L.LEADER=U.USER_CODE where 1=1 "&where&factorywhereoutsideand&" order by L.LINE_NAME"
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
<!--#include virtual="/Language/Job/Shift/Lan_Shift.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Job/Shift/Shift.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="11%" height="20"><span id="inner_SearchLine"></span></td>
    <td width="18%"><input name="linename" type="text" id="linename" value="<%=linename%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_Schedule"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_Records"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
  <td width="241" height="20" class="t-t-Borrow"><div align="center"><span id="inner_LineName"></span></div></td>
  <td width="253" class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?factory='+this.options[this.selectedIndex].value">
      <option value="">Factory</option>
      <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
      <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_Section"></span></div></td>
  <td width="235" height="20" class="t-t-Borrow"><div align="center"><span id="inner_Administrators"></span></div></td>
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
    <td height="20"><div align="center"><%if rs("SHIFT_STATUS")="1" then%><span id="inner_openline" class="red" style="cursor:hand" title="" onClick="javascript:window.open('ShiftIn.asp?id=<%=rs("NID")%>&line_name=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/Start.gif" width="10" height="9"></span><%else%><span id="inner_stopline" class="red" style="cursor:hand" title="" onClick="javascript:window.open('ShiftOut.asp?id=<%=rs("NID")%>&line_name=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/Stop.gif" width="9" height="9"></span><%end if%></div></td>
    <td><div align="center"><span id="inner_openline" class="red" style="cursor:hand" title="" onClick="javascript:window.open('AddSchedule.asp?id=<%=rs("NID")%>&line_name=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconSchedule.gif" width="16" height="20"></span></div></td>
    <td><div align="center"><span id="inner_openline" class="red" style="cursor:hand" title="" onClick="javascript:window.open('ScheduleRecords.asp?id=<%=rs("NID")%>&line_name=<%=rs("LINE_NAME")%>','main')"><img src="/Images/IconRecords.gif" width="16" height="20"></span></div></td>
    <td><div align="center">
      <%if rs("SHIFT_STATUS")="1" then%>Í£Ïß<%else%>¿ªÏß<%end if%>
    </div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('ShiftHistory.asp?linename=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><%= rs("LINE_NAME") %></span></div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
    <td height="20"><div align="center"><%= rs("USER_CHINESE_NAME")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->