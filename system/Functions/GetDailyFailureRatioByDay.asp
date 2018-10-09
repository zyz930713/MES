<%
function GetDailyFailureRatioByDay(seriesgroup_id,defectcode_id,station_id,thisday)
SQL="select RATIO from BAR_REPORT.DAILY_FAILURE_RATIO DY where DY.SERIES_GROUP_ID='"&seriesgroup_id&"' and DY.DEFECTCODE_ID='"&defectcode_id&"' and DY.STATION_ID='"&station_id&"' and to_char(DY.REPORT_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
'response.Write(SQL)
rs.open SQL,conn,1,3
if not rs.eof then
GetDailyFailureRatioByDay=csng(rs("RATIO"))
else
GetDailyFailureRatioByDay=0
end if
rs.close
end function
%>