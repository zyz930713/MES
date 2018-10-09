<%
function getPartTargetYield(part_number_id)
	set rsY=server.CreateObject("adodb.recordset")
	SQLY="Select TARGET_YIELD from Part where NID='"&part_number_id&"'"
	rsY.open SQLY,conn,1,3
	if not rsY.eof then
		if not isnull(rsY("TARGET_YIELD")) then
		getPartTargetYield=rsY("TARGET_YIELD")
		else
		getPartTargetYield=0
		end if
	else
	getPartTargetYield=0
	end if
	rsY.close
	set rsY=nothing
end function
%>