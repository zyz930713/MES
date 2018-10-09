<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
FactoryRight ""
%>
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
<script language="JavaScript" src="/Reports/Yield/DailyStationYield/WeeklyFormcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Process/Failure/WeeklyFailureRatio1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="2" class="t-c-greenCopy"><p>Weekly Failure Ratio Selection</p></td>
    </tr>
    
    <tr>
      <td width="11%" class="t-b-blue"><span>Factory</span></td>
      <td width="89%"><select name="factory_id" id="factory_id">
        <option value="">-- Select --</option>
        <%=getFactory("OPTION","",factorywhereinside,"","")%>
      </select></td>
    </tr>
    <tr>
      <td class="t-b-blue">Job Close Time</td>
      <td>From
        <input name="fromdate" type="text" id="fromdate" value="<%=dateadd("d",-1,date())%>" size="10">
        <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.fromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
        </script>
        <input name="fromhour" type="text" id="fromhour" value="00" size="2">
        :
        <input name="fromminute" type="text" id="fromminute" value="00" size="2">
        &nbsp;to
        <input name="todate" type="text" id="todate" value="<%=date()%>" size="10">
        <script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
		{
		document.all.todate.value=year + '-' + month + '-' + date
		}
		calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
        </script>
        &nbsp;
        <input name="tohour" type="text" id="tohour" value="<%=hour(now())%>" size="2">
        :
        <input name="tominute" type="text" id="tominute" value="<%=minute(now())%>" size="2"></td>
    </tr>
    <tr>
      <td class="t-b-blue">Week Number</td>
      <td><input name="week_number" type="text" id="week_number" size="2" onChange="weeknumbercheck(this)">
        of 
          <input name="year_number" type="text" id="year_number" value="<%=year(date())%>" size="4" onChange="yearnumbercheck(this)"></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><div align="center">
        <input name="path" type="hidden" id="path" value="<%=path%>">
        <input name="query" type="hidden" id="query" value="<%=query%>">
        <input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Generate Weekly Report">
        &nbsp;
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->