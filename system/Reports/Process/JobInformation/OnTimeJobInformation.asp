<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/Process/JobInformation/OnTimeFormcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Process/JobInformation/OnTimeJobInformation1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue">Ontime Job Information Selection </td>
    </tr>
    <tr>
      <td width="11%">Report Name</td>
      <td width="89%"><input name="report_name" type="text" id="report_name" value="<% =session("user")&"'s Job Information Report on "&now()%>" size="100"></td>
    </tr>
    <tr>
      <td>Part Number </td>
      <td><input name="part_number" type="text" id="part_number" size="50"></td>
    </tr>
    <tr>
      <td>Job Number </td>
      <td><input name="job_number" type="text" id="job_number" size="50">
        Please Key-in
        Master Job Number </td>
    </tr>
    <tr>
      <td>Line Name </td>
      <td><input name="line_name" type="text" id="line_name" size="50"></td>
    </tr>
    <tr>
      <td>Job Status </td>
      <td><label>
        <select name="job_status" id="job_status">
          <option value="">-- All --</option>
          <option value="0">Opened</option>
          <option value="1">Closed</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td>Job Start Time</td>
      <td>From
        <input name="start_fromdate" type="text" id="start_fromdate" value="<%=dateadd("d",-1,date())%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.start_fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
        <input name="start_fromhour" type="text" id="start_fromhour" value="00" size="2" onChange="hournumbercheck(this)">
        :
        <input name="start_fromminute" type="text" id="start_fromminute" value="00" size="2" onChange="minutenumbercheck(this)">
        &nbsp;to
<input name="start_todate" type="text" id="start_todate" value="<%=date()%>" size="10">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.start_todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp;
<input name="start_tohour" type="text" id="start_tohour" value="<%=hour(now())%>" size="2" onChange="hournumbercheck(this)">
:
<input name="start_tominute" type="text" id="start_tominute" value="<%=minute(now())%>" size="2" onChange="minutenumbercheck(this)">
<input name="Clear_Start" type="button" id="Clear_Start" value="Clear" onClick="clear_start()"></td>
    </tr>
    <tr>
      <td>Job Close Time</td>
      <td>From
        <input name="close_fromdate" type="text" id="close_fromdate" value="" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.close_fromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
        <input name="close_fromhour" type="text" id="close_fromhour" value="" size="2" onChange="hournumbercheck(this)">
:
<input name="close_fromminute" type="text" id="close_fromminute" value="" size="2" onChange="minutenumbercheck(this)">
&nbsp;to
<input name="close_todate" type="text" id="close_todate" value="" size="10">
<script language=JavaScript type=text/javascript>
function calendar4Callback(date, month, year)
	{
	document.all.close_todate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
                        </script>
&nbsp;
<input name="close_tohour" type="text" id="close_tohour" value="" size="2" onChange="hournumbercheck(this)">
:
<input name="close_tominute" type="text" id="close_tominute" value="" size="2" onChange="minutenumbercheck(this)">
<input name="Clear_Close" type="button" id="Clear_Close" value="Clear" onClick="clear_close()"></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center">
        <input name="path" type="hidden" id="path" value="<%=path%>">
        <input name="query" type="hidden" id="query" value="<%=query%>">
        <input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Generate Email Report">
        &nbsp;</div></td>
    </tr>
  </table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->