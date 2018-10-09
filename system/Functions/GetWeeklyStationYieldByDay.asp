<%
function GetWeeklyStationYieldByDay(year_number,week_number,station_id,from_time,to_time)
SQL="select YIELD,FROM_TIME,TO_TIME from BAR_REPORT.WEEKLY_STATION_YIELD DY where DY.YEAR_NUMBER='"&year_number&"' and DY.WEEK_NUMBER='"&week_number&"' and DY.STATION_ID='"&station_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
GetWeeklyStationYieldByDay=csng(rs("YIELD"))
from_time=rs("FROM_TIME")
to_time=rs("TO_TIME")
else
GetWeeklyStationYieldByDay=0
from_time=null
to_time=null
end if
rs.close
end function

function GetWeeklyStationYieldSpan(year_number,week_number,from_time,to_time)
SQL="select FROM_TIME,TO_TIME from BAR_REPORT.WEEKLY_STATION_YIELD DY where DY.YEAR_NUMBER='"&year_number&"' and DY.WEEK_NUMBER='"&week_number&"'and rownum=1"
rs.open SQL,conn,1,3
if not rs.eof then
from_time=rs("FROM_TIME")
to_time=rs("TO_TIME")
else
from_time=null
to_time=null
end if
rs.close
end function
%>