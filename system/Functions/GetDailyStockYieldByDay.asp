<%
function GetDailyStockYieldByDay(thisday)
SQL="select 'F' as REPORT_TYPE,DY.*,F.FACTORY_NAME,P.TASK_NAME,P.PARAM2 as START_HOUR,P.PARAM3 as END_HOUR from DAILY_FAMILYSTOCK_YIELD DY inner join FACTORY F on DY.FACTORY_ID=f.NID inner join PROFILE_TASK P on DY.PROFILE_TASK_ID=P.NID where DY.FAMILY_NAME='OVERALL' and to_char(DY.LINE_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"' union select 'S' as REPORT_TYPE,DY.*,F.FACTORY_NAME,P.TASK_NAME,P.PARAM2 as START_HOUR,P.PARAM3 as END_HOUR from DAILY_SERIESSTOCK_YIELD DY inner join FACTORY F on DY.FACTORY_ID=f.NID inner join PROFILE_TASK P on DY.PROFILE_TASK_ID=P.NID where DY.SERIES_NAME='OVERALL' and to_char(DY.LINE_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
GetDailyStockYieldByDay="<table>"
if not rs.eof then
	while not rs.eof
	GetDailyStockYieldByDay=GetDailyStockYieldByDay&"<tr><td><span style=""cursor:hand"" onClick=""window.open('/Reports/Yield/DailyStockYield/"
	if rs("REPORT_TYPE")="F" then
	GetDailyStockYieldByDay=GetDailyStockYieldByDay&"DailyStockYieldReport.asp"
	else
	GetDailyStockYieldByDay=GetDailyStockYieldByDay&"DailyStockYieldReport_Series.asp"
	end if
	from_time=dates&" "&rs("START_HOUR")
	to_time=dateadd("d",1,dates)&" "&rs("END_HOUR")
	GetDailyStockYieldByDay=GetDailyStockYieldByDay&"?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&from_time&"&to_time="&to_time&"')"" title="""&rs("TASK_NAME")&" - "&dates&"""><li>"&rs("FACTORY_NAME")&":"&rs("STOCK_OUTPUT_QUANTITY")&"/"&rs("STOCK_INPUT_QUANTITY")&"["&formatpercent(csng(rs("FINAL_YIELD")),2,-1)&"]</span>&nbsp;<span class=""red"" onClick=""window.open('RegenerateDailyStockYieldReport.asp?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&from_time&"&to_time="&to_time&"')"" title=""Regenerate"">R</span>&nbsp;<span class=""red"" onClick=""window.open('DailyStockYieldReport.asp?profile_task_id="&rs("PROFILE_TASK_ID")&"&factory_id="&rs("FACTORY_ID")&"&from_time="&dateadd("d",-weekday(dates),from_time)&"&to_time="&to_time&"')"">WTD</span></td></tr>"
	rs.movenext
	wend
	GetDailyStockYieldByDay=GetDailyStockYieldByDay&"</table>"
else
GetDailyStockYieldByDay="None"
end if
rs.close
end function
%>