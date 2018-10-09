<%
function GetDailyLineWIPByDay(thisday)
SQL="select HWL.*,F.FACTORY_NAME,P.TASK_NAME,P.PARAM2 as START_HOUR,P.PARAM3 as END_HOUR from BAR_REPORT.HOURLY_WIP_LIST HWL inner join BAR_WEB.FACTORY F on HWL.FACTORY_ID=f.NID inner join BAR_WEB.PROFILE_TASK P on HWL.PROFILE_TASK_ID=P.NID where to_char(HWL.REPORT_TIME,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"' and REPORT_TYPE='BY_LINE' ORDER BY REPORT_TIME DESC"
rs.open SQL,conn,1,3
GetDailyLineWIPByDay="<table>"
if not rs.eof then
	while not rs.eof
	GetDailyLineWIPByDay=GetDailyLineWIPByDay&"<tr><td><span style=""cursor:hand"" onClick=""window.open('/Reports/Process/DailyLineWIP/HourlyLineWIPReport.asp?WIP_ID="&rs("NID")&"&WIP_NAME="&rs("WIP_NAME")&"&factory_id="&rs("FACTORY_ID")&"')"" title="""&rs("TASK_NAME")&"""><li>"&rs("FACTORY_NAME")&":["&formatdate(rs("REPORT_TIME"),application("veryshorttimeformat"))&"]</span>&nbsp;</td></tr>"
	'<span class=""red"" onClick=""window.open('/Reports/Process/DailyWIP/HourlyWIPReport.asp?WIP_NAME="&rs("WIP_NAME")&"&factory_id="&rs("FACTORY_ID")&"&from_day="&dateadd("d",-weekday(thisday),thisday)&"&to_day="&thisday&"')"" title=""Week to Today"">WTD</span>
	rs.movenext
	wend
	GetDailyLineWIPByDay=GetDailyLineWIPByDay&"</table>"
else
GetDailyLineWIPByDay="None"
end if
rs.close
end function

function GetDailyFamilyWIPByDay(thisday)
SQL="select HWL.*,F.FACTORY_NAME,P.TASK_NAME,P.PARAM2 as START_HOUR,P.PARAM3 as END_HOUR from BAR_REPORT.HOURLY_WIP_LIST HWL inner join BAR_WEB.FACTORY F on HWL.FACTORY_ID=f.NID inner join BAR_WEB.PROFILE_TASK P on HWL.PROFILE_TASK_ID=P.NID where to_char(HWL.REPORT_TIME,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"' and REPORT_TYPE='BY_FAMILY' ORDER BY REPORT_TIME DESC"
rs.open SQL,conn,1,3
GetDailyFamilyWIPByDay="<table>"
if not rs.eof then
	while not rs.eof
	GetDailyFamilyWIPByDay=GetDailyFamilyWIPByDay&"<tr><td><span style=""cursor:hand"" onClick=""window.open('/Reports/Process/DailyFamilyWIP/HourlyFamilyWIPReport.asp?WIP_ID="&rs("NID")&"&WIP_NAME="&rs("WIP_NAME")&"&factory_id="&rs("FACTORY_ID")&"')"" title="""&rs("TASK_NAME")&"""><li>"&rs("FACTORY_NAME")&":["&formatdate(rs("REPORT_TIME"),application("veryshorttimeformat"))&"]</span>&nbsp;</td></tr>"
	'<span class=""red"" onClick=""window.open('/Reports/Process/DailyWIP/HourlyWIPReport.asp?WIP_NAME="&rs("WIP_NAME")&"&factory_id="&rs("FACTORY_ID")&"&from_day="&dateadd("d",-weekday(thisday),thisday)&"&to_day="&thisday&"')"" title=""Week to Today"">WTD</span>
	rs.movenext
	wend
	GetDailyFamilyWIPByDay=GetDailyFamilyWIPByDay&"</table>"
else
GetDailyFamilyWIPByDay="None"
end if
rs.close
end function
%>