<%
function getLine(showstyle,line,where,order,splitchar)
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'line=selected line
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID "&where&" order by L.LINE_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=line then
			output=output&" selected"
			end if
		output=output&">"&rsS("LINE_NAME")&"</option>"
		case "CMMSOPTION"
		output=output&"<option value='"&rsS("LINE_NAME")&"'"
			if rsS("LINE_NAME")=line then
			output=output&" selected"
			end if
		output=output&">"&rsS("LINE_NAME")&"</option>"
		case "TEXT"
		output=output&rsS("LINE_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getLine=output
rsS.close
set rsS=nothing
end function

function getLine2(showstyle,line,where,order,splitchar)
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'line=selected line
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID "&where&" order by L.LINE_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("LINE_NAME")&"'"
			if rsS("LINE_NAME")=line then
			output=output&" selected"
			end if
		output=output&">"&rsS("LINE_NAME")&"</option>"
		case "CMMSOPTION"
		output=output&"<option value='"&rsS("LINE_NAME")&"'"
			if rsS("LINE_NAME")=line then
			output=output&" selected"
			end if
		output=output&">"&rsS("LINE_NAME")&"</option>"
		case "TEXT"
		output=output&rsS("LINE_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getLine2=output
rsS.close
set rsS=nothing
end function
%>