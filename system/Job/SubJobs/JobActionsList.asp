<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
where=""
if station="" or station="all" then
where=where&""
elseif station="null" then
where=where&" and A.STATION_ID is null"
else
where=where&" and A.STATION_ID='"&station&"'"
end if
if purpose="" or purpose="all" then
where=where&""
else
where=where&" and A.ACTION_PURPOSE="&purpose
end if
pagepara="&station="&station&"&purpose="&purpose
SQL="select J.CONTROL_TYPE,J.CONTROL_STATION,J.CONTROL_REASON,J.CONTROL_PERSON,J.CONTROL_TIME from JOB J where J.JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"&where
rs.open SQL,conn,1,3
if not rs.eof then
CONTROL_TYPE=split(left(rs("CONTROL_TYPE"),len(rs("CONTROL_TYPE"))-1),",")
CONTROL_STATION=split(left(rs("CONTROL_STATION"),len(rs("CONTROL_STATION"))-1),",")
CONTROL_REASON=split(left(rs("CONTROL_REASON"),len(rs("CONTROL_REASON"))-1),",")
CONTROL_PERSON=split(left(rs("CONTROL_PERSON"),len(rs("CONTROL_PERSON"))-1),",")
CONTROL_TIME=split(left(rs("CONTROL_TIME"),len(rs("CONTROL_TIME"))-1),",")
end if
rs.close
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
  <td height="20" colspan="6" class="t-c-greenCopy">Browse Job Actions List (<%=jobnumber%>-<%=repeatstring(sheetnumber,"0",3)%>) </td>
</tr>
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy">User:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="6">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Action Name </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Station</div></td>
  <td class="t-t-Borrow"><div align="center">Person</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Reason</div></td>
  <td class="t-t-Borrow"><div align="center">Time</div></td>
  </tr>
<%for i=0 to ubound(CONTROL_TYPE)%>
<tr>
  <td height="20"><div align="center"><%=i+1%></div></td>
    <td height="20"><div align="center">
	<%select case CONTROL_TYPE(i)
	case "0"
	CTYPE="Start"
	case "2"
	CTYPE="Pause"
	case "3"
	CTYPE="Unlock"
	case "4"
	CTYPE="Abort"
	case "5"
	CTYPE="Repeat"
	case "6"
	CTYPE="Alter"
	end select%><%= CTYPE %></div></td>
    <td height="20"><div align="center"><%= getStation("","TEXT",""," where S.NID='"&CONTROL_STATION(i)&"'","","","") %></div></td>
    <td height="20"><div align="center"><%= getRoleMember("","TEXT",""," where USER_CODE='"&CONTROL_PERSON(i)&"'","","","")%>(<%=CONTROL_PERSON(i)%>)&nbsp;</div></td>
    <td height="20"><div align="center"><%= CONTROL_REASON(i) %></div></td>
    <td><div align="center"><% = formatdate(CONTROL_TIME(i),application("longdateformat")) %>&nbsp;</div></td>
  </tr>
  <%next
  if DBA=true then%>
<tr>
  <td height="20" colspan="6"><div align="center">
    <input type="button" name="Button" value="Clear Actions" onClick="javascript:location.href='ClearJobActions.asp?jobnumber=<%=jobnumber%>'">
  </div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->