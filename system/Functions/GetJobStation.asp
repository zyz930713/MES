<%
function getJobStation(jobnumber,sheetnumber,jobtype,station,repeated_sequence,operatorcode,practised,station_start_time,station_close_time)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
if repeated_sequence="" then
repeated_sequence=1
end if
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JS.OPERATOR_CODE,JS.START_TIME,JS.CLOSE_TIME,O.PRACTISED from JOB_STATIONS JS left join OPERATORS O on JS.OPERATOR_CODE=O.CODE where JS.JOB_NUMBER='"&jobnumber&"' and JS.SHEET_NUMBER='"&sheetnumber&"' and JS.JOB_TYPE='"&jobtype&"' and JS.STATION_ID='"&station&"' and JS.REPEATED_SEQUENCE='"&repeated_sequence&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
operatorcode=rsS("OPERATOR_CODE")
practised=rsS("PRACTISED")
station_start_time=rsS("START_TIME")
station_close_time=rsS("CLOSE_TIME")
end if
rsS.close
set rsS=nothing
end function

function getJobStation_History(jobnumber,sheetnumber,jobtype,station,repeated_sequence,operatorcode,practised,station_start_time,station_close_time)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
if repeated_sequence="" then
repeated_sequence=1
end if
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JS.OPERATOR_CODE,JS.START_TIME,JS.CLOSE_TIME,O.PRACTISED from  bar_hist.JOB_STATIONS JS left join OPERATORS O on JS.OPERATOR_CODE=O.CODE where JS.JOB_NUMBER='"&jobnumber&"' and JS.SHEET_NUMBER='"&sheetnumber&"' and JS.JOB_TYPE='"&jobtype&"' and JS.STATION_ID='"&station&"' and JS.REPEATED_SEQUENCE='"&repeated_sequence&"'"

 
rsS.open SQLS,conn,1,3
if not rsS.eof then
operatorcode=rsS("OPERATOR_CODE")
practised=rsS("PRACTISED")
station_start_time=rsS("START_TIME")
station_close_time=rsS("CLOSE_TIME")
end if
rsS.close
set rsS=nothing
end function
%>