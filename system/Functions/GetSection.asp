<%
function getSection(showstyle,section,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select S.NID,S.SECTION_NAME from SECTION S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=section then
			output=output&"selected"
			end if
		output=output&">"&rsP("SECTION_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='section' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("SECTION_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("SECTION_NAME")&splitchar
		end select
	rsP.movenext
	wend
end if
getSection=output
rsP.close
set rsP=nothing
end function
%>