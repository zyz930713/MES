<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Reports/Tracking/JobSchedule/ScheduleList.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<form name="form1" enctype="multipart/form-data" method="post" action="/Devices/PC/ImportPC1.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy">Import Desktop </td>
  </tr>
  <tr>
    <td width="13%" height="20">Desktop File </td>
    <td width="87%" height="20">
      <input type="file" name="file">    </td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
    </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input type="submit" name="Submit" value="Import">
      &nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->