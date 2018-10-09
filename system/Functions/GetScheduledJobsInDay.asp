<%
function getScheduledJobsInDay(schedule_id,part_number,schedule_day)	
getScheduledJobsInDay=0
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select JOB_NUMBER from JOB_SCHEDULE_DETAIL where SCHEDULE_ID='"&schedule_id&"' and PART_NUMBER_ID='"&part_number&"' and STATUS=1 and (extract(year from START_TIME)='"&year(schedule_day)&"' and extract(month from START_TIME)='"&month(schedule_day)&"' and extract(day from START_TIME)='"&day(schedule_day)&"')"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getScheduledJobsInDay=rsS.recordcount*10
end if
rsS.close
set rsS=nothing
end function
%>