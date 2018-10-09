<%
if session("role")<>"" then
	if instr(session("role"),",TESTDEFECTCODE_ADMINISTRATOR")<>0 then
	admin=true
	else
		admin=false
		if instr(session("role"),",TESTDEFECTCODE_READER")=0 then
		response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>