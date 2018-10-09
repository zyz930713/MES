<%
function getJobActionValue(jobnumber,sheetnumber,jobtype,station,action)	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select ACTION_VALUE from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' and STATION_ID='"&station&"' and ACTION_ID='"&action&"'"
session("SQLS")=SQLS
rsS.open SQLS,conn,1,3
if not rsS.eof then
getJobActionValue=rsS("ACTION_VALUE")
else
getJobActionValue=""
end if
rsS.close
set rsS=nothing
end function
%>