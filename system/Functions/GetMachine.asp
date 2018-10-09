<%
function getMachine(isshowname,showstyle,machine,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsM=server.CreateObject("adodb.recordset")
SQLM="Select NID,MACHINE_NUMBER,MACHINE_NAME from MACHINE "&where&order
rsM.open SQLM,conn,1,3
if not rsM.eof then
while not rsM.eof
	select case showstyle
	case "OPTION"
	output=output&"<option value='"&rsM("NID")&"'"
		if rsM("NID")=machine then
		output=output&"selected"
		end if
		if isshowname=true then
		output=output&">"&rsM("MACHINE_NUMBER")&" ("&rsM("MACHINE_NAME")&")"&"</option>"
		else
		output=output&">"&rsM("MACHINE_NAME")&"</option>"
		end if
	case "TEXT"
	output=output&rsM("MACHINE_NAME")&splitchar
	end select
rsM.movenext
wend
end if
getMachine=output
rsM.close
set rsM=nothing
end function
%>