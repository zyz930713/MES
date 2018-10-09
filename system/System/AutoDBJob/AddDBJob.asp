<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/AutoDBJob/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<!--#include virtual="/Language/System/AutoDBJob/Lan_AddDBJob.asp" -->
</head>

<body onLoad="language();">
<form action="/System/AutoDBJob/AddDBJob1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span> <%= line_name %></td>
</tr>
  <tr>
    <td height="20"><span id="inner_Administrator"></span></td>
    <td height="20"><% = session("code") %>&nbsp;</td>
  </tr>
<tr>
  <td width="102" height="20"><span id="inner_JobItem"></span></td>
  <td width="820" height="20" class="red"><select name="jobitem" id="jobitem">
    <option value="JOB_MASTER_SCRAP_UPDATE$">JOB_MASTER_SCRAP_UPDATE$</option>
  </select></td>
    </tr>
  <tr>
    <td height="20"><span id="inner_ScheduleDay"></span> <span class="red">*</span></td>
    <td height="20"><label>
      <select name="period" id="period">
        <option value="">--Select Period--</option>
        <option value="sysdate+1">Each Day</option>
		<option value="sysdate+2">Each Two Days</option>
		<option value="sysdate+7">Weekly</option>
      </select>
    </label></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_ScheduleTime"></span> <span class="red">*</span> </td>
    <td height="20"><input name="day" type="text" id="day" value="<%=date()%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.day.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
      </script>
      <select name="hour" id="hour">
	<option value="">-- Select --</option>
	<%for i=0 to 24%>
	<option value="<%=i%>" <%if i=hour(now) then%>selected<%end if%>><%=i%></option>
	<%next%>
    </select>
	<span id="inner_Hour"></span>
	<select name="minute" id="minute">
	<option value="">-- Select --</option>
	<%for i=0 to 59%>
	<option value="<%=i%>"<%if i=minute(now) then%>selected<%end if%>><%=i%></option>
	<%next%>
     </select>
	<span id="inner_Minute"></span></td>
  </tr>
  
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="line_id" type="hidden" id="line_id" value="<%=line_id%>">
      <input name="line_name" type="hidden" id="line_name" value="<%=line_name%>">
	  <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input name="Save" type="submit" id="Save" value="Save">
&nbsp;
<input name="Reset" type="reset" id="Reset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->