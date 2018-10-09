<%
if session("role")<>"" then
	if instr(session("role"),",SYSTEM_ADMINISTRATOR")=0 then
	response.Redirect("/main.asp?error=Error message: you are not system administrator and can not access this page.")
	end if
else
response.Redirect("/main.asp?error=Error message: you are not system administrator and can not access this page.")
end if
%>