<%
function GetDailyStationYieldByDay(station_id,thisday)
SQL="select YIELD from BAR_REPORT.DAILY_STATION_YIELD DY where DY.STATION_ID='"&station_id&"' and to_char(DY.REPORT_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
if not rs.eof then
GetDailyStationYieldByDay=csng(rs("YIELD"))
else
GetDailyStationYieldByDay=0
end if
rs.close
end function
%>