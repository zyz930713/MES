<%
function getJobStationGoodQuantity(jobnumber,sheetnumber,station)	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select GOOD_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and STATION_ID='"&station&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getJobStationGoodQuantity=rsS("GOOD_QUANTITY")
else
getJobStationGoodQuantity=0
end if
rsS.close
set rsS=nothing
end function
%>