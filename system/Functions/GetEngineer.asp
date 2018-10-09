<%
function getEngineer(showstyle,engineer,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select U.USER_CODE,U.USER_NAME,U.USER_CHINESE_NAME,U.EMAIL from USERS U "&where&order
rsP.open SQLP,conn,1,3
response.Write(SQLP)
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("USER_CODE")&"'"
			if rsP("USER_CODE")=engineer then
			output=output&"selected"
			end if
		output=output&">"&rsP("USER_CODE")&" - "&rsP("USER_CHINESE_NAME")&" ("&rsP("USER_NAME")&")</option>"
		case "RADIO"
		output=output&"<input name='section' type='radio' class='t-c-greenCopy' value='"&rsP("USER_CODE")&"'>"&rsP("USER_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("USER_NAME")&splitchar
		end select
	rsP.movenext
	wend
end if
getEngineer=output
rsP.close
set rsP=nothing
end function
%>