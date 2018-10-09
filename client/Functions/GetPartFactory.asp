<%
function getPartFactory(part)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select FACTORY_ID from PART where NID='"&part&"'"
rsP.open SQLP,conn,1,3
if not rsP.eof then
getPartFactory=rsP("FACTORY_ID")
end if
rsP.close
set rsP=nothing
end function
%>