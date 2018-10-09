<%
function getFactory(showstyle,factory,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select NID,FACTORY_NAME from FACTORY "&where&order
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
		output=output&">"&rsP("FACTORY_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='factory' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("FACTORY_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("FACTORY_NAME")&splitchar
		end select
	rsP.movenext
	wend
end if
getFactory=output
rsP.close
set rsP=nothing
end function
%>