<%
function getLine(showstyle,line,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=line then
			output=output&"selected"
			end if
		output=output&">"&rsP("LINE_NAME")&"</option>"
		case "OPTION_LINE_NAME"
		output=output&"<option value='"&rsP("LINE_NAME")&"'"
			if rsP("LINE_NAME")=line then
			output=output&"selected"
			end if
		output=output&">"&rsP("LINE_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='factory' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("LINE_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("LINE_NAME")&splitchar
		end select
	rsP.movenext
	wend
end if
getLine=output
rsP.close
set rsP=nothing
end function
%>