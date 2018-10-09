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
<script language="JavaScript" src="/Reports/Cycle/FamilyCycle/OnTimeFormcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Cycle/FamilyCycle/OnTimeFamilyCycle1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue">Ontime Family Cycle Time Selection </td>
    </tr>
    <tr>
      <td width="11%">Report Name</td>
      <td width="89%"><input name="report_name" type="text" id="report_name" value="<% =session("user")&"'s Family Cycle Report on "&now()%>" size="100"></td>
    </tr>
    <tr>
      <td>Job Close Time</td>
      <td>From
        <input name="close_fromdate" type="text" id="close_fromdate" value="<%=dateadd("d",-1,date())%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.close_fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
        <input name="close_fromhour" type="text" id="close_fromhour" value="00" size="2" onChange="hournumbercheck(this)">
        :
        <input name="close_fromminute" type="text" id="close_fromminute" value="00" size="2" onChange="minutenumbercheck(this)">
        &nbsp;to
<input name="close_todate" type="text" id="close_todate" value="<%=date()%>" size="10">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.close_todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp;
<input name="close_tohour" type="text" id="close_tohour" value="<%=hour(now())%>" size="2" onChange="hournumbercheck(this)">
:
<input name="close_tominute" type="text" id="close_tominute" value="<%=minute(now())%>" size="2" onChange="minutenumbercheck(this)">
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