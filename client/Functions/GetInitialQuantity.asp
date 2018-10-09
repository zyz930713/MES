<%
function getInitialQuantity(jobnumber,sheetnumber,jobtype,station)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	'RESPONSE.WRITE JOBNUMBER &"-"&SHEETNUMBER&"-"&JOBTYPE&"-"
	set rsS=server.CreateObject("adodb.recordset")
 
	SQLS="Select GOOD_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' and STATION_ID='"&station&"'"
 
rsS.open SQLS,conn,1,3
if not rsS.eof then
good_quantity=rsS("GOOD_QUANTITY")
end if
rsS.close
set rsS=nothing
getInitialQuantity=good_quantity
end function
%>