<%
function getJobStation(jobnumber,sheetnumber,jobtype,station,operatorcode,station_start_time,station_close_time)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select OPERATOR_CODE,START_TIME,nvl(CLOSE_TIME,SYSDATE) as CLOSE_TIME from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' and STATION_ID='"&station&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
operatorcode=rsS("OPERATOR_CODE")
station_start_time=rsS("START_TIME")
station_close_time=now()
end if
rsS.close
set rsS=nothing
end function
%>