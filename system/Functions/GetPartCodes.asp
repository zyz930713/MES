<%
function getPartCodes(part_number,splitchar,idcount)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsO=server.CreateObject("adodb.recordset")
SQLO="Select CODE from OPERATORS where AUTHORIZED_PARTS_ID like '%"&part_number&"%' order by CODE"
rsO.open SQLO,conn,1,3
idcount=rsO.recordcount
z=rsO.recordcount
if not rsO.eof then
y=1
while not rsO.eof
	if y<>z then
	output=output&rsO("CODE")&splitchar
	else
	output=output&rsO("CODE")
	end if
rsO.movenext
y=y+1
wend
end if
getPartCodes=output
rsO.close
set rsO=nothing
end function
%>