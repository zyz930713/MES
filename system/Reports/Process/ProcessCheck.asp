<%
if instr(session("role"),",PROCESS_ADMINISTRATOR")<>0 then
admin=true
else
	admin=false
	if instr(session("role"),",PROCESS_READER")=0 then
	response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
	end if
end if
%>