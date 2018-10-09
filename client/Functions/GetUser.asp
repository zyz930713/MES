<%
function getUser(showstyle,user,where,order,splitchar)
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'line=selected line
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select U.USER_CODE,U.USER_NAME,U.USER_CHINESE_NAME from USERS U "&where&" order by U.USER_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("USER_CODE")&"'"
			if rsS("USER_CODE")=user then
			output=output&" selected"
			end if
		output=output&">"&rsS("USER_NAME")&" - "&rsS("USER_CHINESE_NAME")&"</option>"
		case "TEXT"
		if rsS("USER_CODE")=user then
		output=output&rsS("USER_NAME")&splitchar
		end if
		end select
	rsS.movenext
	wend
end if
getUser=output
rsS.close
set rsS=nothing
end function
%>