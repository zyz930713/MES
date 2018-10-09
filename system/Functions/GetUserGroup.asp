<%
function getUserGroup(showstyle,usergroup,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select U.NID,U.GROUP_NAME,U.GROUP_CHINESE_NAME from USER_GROUP U inner join FACTORY F on U.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=usergroup then
			output=output&"selected"
			end if
		output=output&">"&rsP("GROUP_NAME")&" - "&rsP("GROUP_CHINESE_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='section' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("GROUP_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("GROUP_NAME")&splitchar
		case "TEXT_NID"
			output=output&rsP("NID")&splitchar
		end select
	rsP.movenext
	wend
end if
getUserGroup=output
rsP.close
set rsP=nothing
end function
%>