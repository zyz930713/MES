<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetDailyFinanceYieldByDay.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetProfileTask.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")

if request.QueryString("yieldmonth")<>"" then
thisyieldmonth=cint(request.QueryString("yieldmonth"))
else
thisyieldmonth=month(date())
end if

if request.QueryString("yieldyear")<>"" then
thisyieldyear=cint(request.QueryString("yieldyear"))
else
thisyieldyear=year(date())
end if
pagename="/Reports/Yield/DailyFinanceYield/DailyFinanceYieldList.asp"
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
<script language="JavaScript" src="/Reports/Yield/DailyFinanceYield/formcheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Reports/Yield/DailyFinanceYield/DailyFinanceYieldReport.asp" method="post" name="form1" target="_blank" onSubmit="return formcheck()">
  <!--<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#DFDFDF">
    <tr>
      <td height="20" colspan="7" class="t-c-greenCopy">Detail Span Selection </td>
    </tr>
    <tr>
      <td class="t-b-blue">Task Type </td>
      <td><select name="profile_task_id" id="profile_task_id">
          <option value="">--Select--</option>
          <%'= getProfileTask("OPTION",""," and (lower(TASK_NAME) like '%daily family finance yield%' or lower(TASK_NAME) like '%daily series line yield%')"," order by TASK_NAME","") %>
      </select></td>
      <td class="t-b-blue">Factory: </td>
      <td><select name="factory_id" id="factory_id">
          <option value="">-- Select --</option>
          <%'=getFactory("OPTION","",factorywhereinside,"","")%>
      </select></td>
      <td class="t-b-blue">Job Close Time</td>
      <td>From
        <input name="fromdate" type="text" id="fromdate" value="<%'=dateadd("d",-6,date())%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
        <input name="fromhour" type="text" id="fromhour" value="00:00" size="5">
&nbsp;to
<input name="todate" type="text" id="todate" value="<%'=date()%>" size="10">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp;
<input name="tohour" type="text" id="tohour" value="00:00" size="5"></td>
      <td><input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Multi-Day Report"></td>
    </tr>
  </table>-->
</form>
<table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" byield="1" byieldcolorlight="#666666" byieldcolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">
      <table width="100%" cellpadding="0" cellspacing="0" byield="0">
        <tr>
          <td>&nbsp;
              <%
			  if thisyieldmonth=12 then
			  nextyieldyear=thisyieldyear+1
			  nextyieldmonth=1
			  preyieldyear=thisyieldyear
			  preyieldmonth=thisyieldmonth-1
			  elseif thisyieldmonth=1 then
			  preyieldyear=thisyieldyear-1
			  preyieldmonth=12
			  nextyieldyear=thisyieldyear
			  nextyieldmonth=thisyieldmonth+1
			  else
			  preyieldyear=thisyieldyear
			  preyieldmonth=thisyieldmonth-1
			  nextyieldyear=thisyieldyear
			  nextyieldmonth=thisyieldmonth+1
			  end if
			  %>
		  <a href="<%=pagename%>?yieldyear=<%=preyieldyear%>&yieldmonth=<%=preyieldmonth%>" class="white">Previous Month</a></td>
          <td><div align="center">Daily Finance Yield of 
              <%=monthconvert(thisyieldmonth)%></div></td>
          <td><div align="right"><a href="<%=pagename%>?yieldyear=<%=nextyieldyear%>&yieldmonth=<%=nextyieldmonth%>" class="white">Next Month</a></div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="t-b-blue">
    <td height="20" class="red"><div align="center">Sunday</div></td>
    <td height="20"><div align="center">Monday</div></td>
    <td height="20"><div align="center">Tuesday</div></td>
    <td height="20"><div align="center">Wednesday</div></td>
    <td height="20"><div align="center">Thursday</div></td>
    <td height="20"><div align="center">Friday</div></td>
    <td height="20" class="green"><div align="center">Saturday</div></td>
  </tr>
  <%
		dates=cdate(thisyieldyear&"-"&thisyieldmonth&"-1")
		While (year(dates)=thisyieldyear and month(dates)=thisyieldmonth)
		%>
  <tr>
    <%		bgimg=""
			yields=""
			i=1
			while i<=7
			if dates=date() then
			currentclass="t-b-blue"
			else
			currentclass=""
			end if
			if weekday(dates)=i and year(dates)=thisyieldyear and month(dates)=thisyieldmonth then
				yields=GetDailyFinanceYieldByDay(dates)
				if i=1 then
				bgimg="Red_"&day(dates)
				elseif i=7 then
				bgimg="Green_"&day(dates)
				else
				bgimg="Gray_"&day(dates)
				end if
			dates=dateadd("d","1",dates)
			else
			yields="&nbsp;"
			end if
	%>
    <td width="110" height="110" valign="middle" background="/Images/Date/<%=bgimg%>.gif" class="<%=currentclass%>" ><div align="center"><%=yields%></div></td>
    <%	
			i=i+1
			bgimg=""
			yields=""
			wend
	%>
  </tr>
  	<%
		wend
	%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->