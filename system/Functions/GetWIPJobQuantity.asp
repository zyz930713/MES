<%
function getWIPJobQuantity(jobnumber,sheetnumber,jobtype,station_id,merged_stations_id)	
getWIPJobGoodQuantity=0
	if merged_stations_id<>"" then
	a_merged_stations_id=split(merged_stations_id,",")
		for p=0 to ubound(a_merged_stations_id)
		where=where&" or J.CURRENT_STATION_ID='"&a_merged_stations_id(p)&"'"
		next
	end if
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JA.ACTION_VALUE from JOB_ACTIONS JA inner join JOB J on JA.JOB_NUMBER=J.JOB_NUMBER and JA.SHEET_NUMBER=J.SHEET_NUMBER where (J.CURRENT_STATION_ID='"&station_id&"' "&where&") and JA.JOB_NUMBER='"&jobnumber&"' and JA.SHEET_NUMBER='"&sheetnumber&"' and JA.JOB_TYPE='"&jobtype&"' and J.STATUS=0 and JA.STATION_ID=J.FIRST_STATION_ID and JA.ACTION_ID='AC00000097'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getWIPJobQuantity=rsS("ACTION_VALUE")
end if
rsS.close
set rsS=nothing
end function
%>