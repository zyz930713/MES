<%
function getMaterial(isshowname,showstyle,material,where,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsM=server.CreateObject("adodb.recordset")
SQLM="Select 1,M.NID,M.MATERIAL_NAME from MATERIAL M inner join FACTORY F on M.FACTORY_ID=F.NID "& where
rsM.open SQLM,conn,1,3
if not rsM.eof then
while not rsM.eof
	select case showstyle
	case "OPTION"
	output=output&"<option value='"&rsM("NID")&"'>"&rsM("MATERIAL_NAME")&"</option>"
	case "TEXT"
	output=output&rsM("MATERIAL_NAME")&splitchar
	end select
rsM.movenext
wend
end if
getMaterial=output
rsM.close
set rsM=nothing
end function
%>