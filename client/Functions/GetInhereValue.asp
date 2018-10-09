<%
function getInhereValue(action_id,machine)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select ACTION_VALUE from ACTION_INHERE where ACTION_ID='"&action_id&"' and MACHINE='"&machine&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getInhereValue=rsS("ACTION_VALUE")
end if
rsS.close
set rsS=nothing
end function
%>