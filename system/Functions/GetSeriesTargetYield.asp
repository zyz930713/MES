<%
function getSeriesTargetYield(series_id)
	set rsY=server.CreateObject("adodb.recordset")
	SQLY="Select TARGET_YIELD from SERIES where NID='"&series_id&"'"
	rsY.open SQLY,conn,1,3
	if not rsY.eof then
		if not isnull(rsY("TARGET_YIELD")) then
		getSeriesTargetYield=rsY("TARGET_YIELD")
		else
		getSeriesTargetYield=0
		end if
	else
	getSeriesTargetYield=0
	end if
	rsY.close
	set rsY=nothing
end function
%>