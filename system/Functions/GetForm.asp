<%
function getForm(showstyle,thisform,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select NID,FORM_NAME,FORM_CHINESE_NAME from FORM "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=thisform then
			output=output&"selected"
			end if
		output=output&">"&rsP("FORM_NAME")&" - "&rsP("FORM_CHINESE_NAME")&"</option>"
		end select
	rsP.movenext
	wend
end if
getForm=output
rsP.close
set rsP=nothing
end function
%>