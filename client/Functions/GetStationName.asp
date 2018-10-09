<%
function getStationName(isshowname,station,chinesename)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where NID='"&station&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
output=rsS("STATION_NAME")
chinesename=rsS("STATION_CHINESE_NAME")
end if
getStationName=output
rsS.close
set rsS=nothing
end function
%>