<%
function getApprovalRole(showstyle,role,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsE=server.CreateObject("adodb.recordset")
SQLE="Select NID,ROLE_NAME from SYSTEM_APPROVAL_ROLE "& where&order
rsE.open SQLE,conn,1,3
if not rsE.eof then
while not rsE.eof
	select case showstyle
	case "OPTION"
	output=output&"<option value='"&rsE("NID")&"'"
		if rsE("NID")=role then
		output=output&"selected"
		end if
	output=output&">"&rsE("ROLE_NAME")&"</option>"
	case "TEXT"
		output=output&rsE("ROLE_NAME")&splitchar
	end select
rsE.movenext
wend
end if
getApprovalRole=output
rsE.close
set rsE=nothing
end function
%>