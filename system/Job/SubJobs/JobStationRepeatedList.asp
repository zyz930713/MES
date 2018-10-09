<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
jobnumber=request.QueryString("jobnumber")
SQL="select JAR.*,S.STATION_NAME,A.ACTION_NAME from JOB_ACTIONS_REPEATED JAR inner join STATION S on JAR.STATION_ID=S.NID inner join ACTION A on JAR.ACTION_ID=A.NID where JAR.JOB_NUMBER='"&jobnumber&"' order by JAR.RECORDED_TIME"
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

<form action="/Job/SubJobs/JobStationRepeatedList1.asp" method="post" name="checkform" id="checkform">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy">Browse Job Repeated Actions List (<%=jobnumber%>) </td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy">User:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="7">&nbsp;</td>
</tr>
<tr>
  <td class="t-t-Borrow"><div align="center">Delete</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Index</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Station</div></td>
  <td class="t-t-Borrow"><div align="center">Action Name </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Action Value </div></td>
  <td class="t-t-Borrow"><div align="center">Operator</div></td>
  <td class="t-t-Borrow"><div align="center">Recorded Time </div></td>
</tr>
<%if not rs.eof then
i=1
while not rs.eof%>
<tr>
  <td>
    <div align="center">
      <input name="station_id<%=i%>" type="hidden" id="station_id<%=i%>" value="<%=rs("STATION_ID")%>">
      <input name="action_id<%=i%>" type="hidden" id="action_id<%=i%>" value="<%=rs("ACTION_ID")%>">
      <input name="repeated_sequence<%=i%>" type="hidden" id="repeated_sequence<%=i%>" value="<%=rs("REPEATED_SEQUENCE")%>">
      <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">
      </div>
    </td>
  <td height="20"><div align="center"><%= rs("REPEATED_SEQUENCE")%></div></td>
    <td height="20"><div align="center"><%= rs("STATION_NAME")%></div></td>
    <td height="20"><div align="center"><%= rs("ACTION_NAME")%></div></td>
    <td height="20"><div align="center"><%= rs("ACTION_VALUE")%></div></td>
    <td><div align="center"><%= rs("OPERATOR_CODE")%></div></td>
    <td><div align="center"><%= rs("RECORDED_TIME")%></div></td>
</tr>
<%
i=i+1
rs.movenext
wend
%>
<tr>
  <td height="20" colspan="7"><div align="center">
  <input name="jobnumber" type="hidden" id="jobnumber" value="<%=jobnumber%>">
  <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
  <input name="path" type="hidden" id="path" value="<%=path%>">
  <input name="query" type="hidden" id="query" value="<%=query%>">
  <input name="Button1" type="button" id="Button1" value="Check All" onClick="checkall()">
  &nbsp;
  <input name="Button2" type="button" id="Button2" value="Uncheck All" onClick="uncheckall()">
  &nbsp;
  <input type="submit" name="Submit" value="Delete">
  &nbsp;
  <input name="Reset" type="reset" id="Reset" value="Reset">
  </div></td>
</tr>
<%else%>
<tr>
  <td height="20" colspan="7"><div align="center">No Records </div>    </td>
  </tr>
<%
end if
rs.close%>
</table>
</form>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->