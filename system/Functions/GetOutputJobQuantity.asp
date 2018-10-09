<%
function getOutputJobQuantity(fromtime,totime,jobnumber,sheetnumber,jobtype,station_id)	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JA.ACTION_VALUE from JOB_STATIONS JS inner join JOB_ACTIONS JA on JS.JOB_NUMBER=JA.JOB_NUMBER and JS.SHEET_NUMBER=JA.SHEET_NUMBER where JA.JOB_NUMBER='"&jobnumber&"' and JA.SHEET_NUMBER='"&sheetnumber&"' and JA.JOB_TYPE='"&jobtype&"' and JS.STATION_ID='"&station_id&"' and JS.STATUS=2 and JS.CLOSE_TIME>=to_date('"&fromtime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME<=to_date('"&totime&"','yyyy-mm-dd hh24:mi:ss') and JA.ACTION_ID='AC00000097'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getOutputJobQuantity=rsS("ACTION_VALUE")
else
getOutputJobQuantity=0
end if
rsS.close
set rsS=nothing
end function
%>