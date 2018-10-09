<%
function getFormUser(showstyle,where,user_code,user_name,email)
set rsU=server.CreateObject("adodb.recordset")
if where<>"" then
SQLU="Select NID,USER_CODE,USER_NAME,EMAIL from USERS "&where
else
SQLU="Select NID,USER_CODE,USER_NAME,EMAIL from USERS where USER_CODE='"&user_code&"'"
end if
'response.Write(SQLU)
'response.End()
rsU.open SQLU,conn,1,3
if not rsU.eof then
	select case showstyle
	case "FORMUSER"
	user_code=rsU("USER_CODE")
	user_name=rsU("USER_NAME")
	email=trim(rsU("EMAIL"))
	end select
end if
rsU.close
set rsU=nothing
end function
%>