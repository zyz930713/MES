<%
function getPartType(showstyle,factory,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select NID,PART_TYPE from Routing "&where&order
'response.Write(SQLP)
'response.End()
rsP.open SQLP,conn,1,3
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=factory then
			output=output&"selected"
			end if
		output=output&">"&rsP("PART_TYPE")&"</option>"
		case "RADIO"
		output=output&"<input name='PART_TYPE' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("PART_TYPE")&"&nbsp;"
		case "TEXT"
			output=output&rsP("PART_TYPE")&splitchar
		end select
	rsP.movenext
	wend
end if
getPartType=output
rsP.close
set rsP=nothing
end function
%>