<%
function getRecordedWIP(WIP_id,line_id,station_id)	
getRecordedWIP=0
SQLS="Select QUANTITY from BAR_REPORT.HOURLY_WIP_DETAIL where WIP_ID='"&WIP_id&"' and LINE_ID='"&line_id&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedWIP=rsS("QUANTITY")
end if
rsS.close
end function

function getRecordedFamilyWIP(WIP_id,family_id,station_id)	
getRecordedFamilyWIP=0
SQLS="Select QUANTITY from BAR_REPORT.HOURLY_WIP_DETAIL where WIP_ID='"&WIP_id&"' and FAMILY_ID='"&family_id&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedFamilyWIP=rsS("QUANTITY")
end if
rsS.close
end function

function getRecordedWIP_WTD(line_id,station_id,job_numbers,from_day,to_day)	
getRecordedWIP_WTD=0
SQLS="Select nvl(avg(QUANTITY),0) as QUANTITY from BAR_REPORT.HOURLY_WIP_DETAIL WD inner join BAR_REPORT.HOURLY_WIP_LIST WL on WD.WIP_ID=WL.NID where WL.REPORT_TYPE='BY_LINE' and to_char(WL.REPORT_TIME,'yyyy-mm-dd')>='"&formatdate(from_day,application("F_shortdateformat"))&"' and to_char(WL.REPORT_TIME,'yyyy-mm-dd')<='"&formatdate(to_day,application("F_shortdateformat"))&"' and LINE_ID='"&line_id&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedWIP_WTD=round(rsS("QUANTITY"),0)
end if
rsS.close
end function

function getRecordedFamilyWIP_WTD(family_id,station_id,job_numbers,from_day,to_day)	
getRecordedFamilyWIP_WTD=0
SQLS="Select nvl(avg(QUANTITY),0) as QUANTITY from BAR_REPORT.HOURLY_WIP_DETAIL WD inner join BAR_REPORT.HOURLY_WIP_LIST WL on WD.WIP_ID=WL.NID where WL.REPORT_TYPE='BY_FAMILY' and to_char(WL.REPORT_TIME,'yyyy-mm-dd')>='"&formatdate(from_day,application("F_shortdateformat"))&"' and to_char(WL.REPORT_TIME,'yyyy-mm-dd')<='"&formatdate(to_day,application("F_shortdateformat"))&"' and FAMILY_ID='"&family_id&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedFamilyWIP_WTD=round(rsS("QUANTITY"),0)
end if
rsS.close
end function

function getRecordedFamilyCapacity(WIP_id,family_id,WIP_target)	
getRecordedFamilyCapacity=0
WIP_target=0
SQLS="Select WIP_TIME,WIP_TARGET from BAR_REPORT.HOURLY_WIP_CAPACITY WD where WD.WIP_ID='"&WIP_id&"' and FAMILY_ID='"&family_id&"'"
'response.Write(SQLS)
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedFamilyCapacity=formatnumber(csng(rsS("WIP_TIME")),2,-1)
WIP_target=round(csng(rsS("WIP_TARGET"))/24/60,2)
end if
rsS.close
end function
%>