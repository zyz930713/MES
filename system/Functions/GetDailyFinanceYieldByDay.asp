<%
function GetDailyFinanceYieldByDay(thisday)
if admin=true then
page="DailyFinanceYieldReportAD.asp"
else
page="DailyFinanceYieldReport.asp"
end if
SQL="select DY.*,F.FACTORY_NAME,P.TASK_NAME,P.PARAM2 as START_HOUR,P.PARAM3 as END_HOUR from DAILY_FINANCE_YIELD DY inner join FACTORY F on DY.FACTORY_ID=f.NID inner join PROFILE_TASK P on DY.PROFILE_TASK_ID=P.NID where DY.FAMILY_NAME='OVERALL' and to_char(DY.LINE_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
GetDailyFinanceYieldByDay="<table>"
if not rs.eof then
	while not rs.eof
	GetDailyFinanceYieldByDay=GetDailyFinanceYieldByDay&"<tr><td><span style=""cursor:hand"" onClick=""window.open('/Reports/Yield/DailyFinanceYield/"&page
	from_time=dates&" "&rs("START_HOUR")
	to_time=dateadd("d",1,dates)&" "&rs("END_HOUR")
	GetDailyFinanceYieldByDay=GetDailyFinanceYieldByDay&"?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&from_time&"&to_time="&to_time&"')"" title="""&rs("TASK_NAME")&" - "&dates&"""><li>"&rs("FACTORY_NAME")&":["&formatpercent(csng(rs("AMOUNT_YIELD")),2,-1)&"]</span>&nbsp;<span class=""red"" onClick=""window.open('RegenerateDailyFinanceYieldReport.asp?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&from_time&"&to_time="&to_time&"')"" title=""Regenerate"">R</span>&nbsp;<span class=""red"" onClick=""window.open('DailyFinanceYieldReport.asp?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&dateadd("d",-weekday(dates),from_time)&"&to_time="&to_time&"')"" title=""Week to Today"">WTD</span></td></tr>"
	rs.movenext
	wend
	GetDailyFinanceYieldByDay=GetDailyFinanceYieldByDay&"</table>"
else
GetDailyFinanceYieldByDay="None"
end if
rs.close
end function
%>