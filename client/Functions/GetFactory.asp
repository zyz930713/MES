<%
function getFactory(showstyle,factory)
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'line=selected line
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select NID,FACTORY_NAME from FACTORY order by FACTORY_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("FACTORY_NAME")=factory then
			output=output&" selected"
			end if
		output=output&">"&rsS("FACTORY_NAME")&"</option>"
		case "NCMROPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("FACTORY_NAME")=factory then
			output=output&" selected"
			end if
		output=output&">"&rsS("FACTORY_NAME")&"</option>"
		end select
	rsS.movenext
	wend
end if
getFactory=output
rsS.close
set rsS=nothing
end function
%>