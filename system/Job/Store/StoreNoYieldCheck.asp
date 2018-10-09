<%
if instr(session("role"),",STORE_ADMINISTRATOR")<>0 then
admin=true
NoYieldchecker=false
	if instr(session("role"),",STORENOYIELD_CHECKER")<>0 then
	NoYieldchecker=true
	end if
else
	if instr(session("role"),",STORENOYIELD_CHECKER")<>0 then
	admin=false
	NoYieldchecker=true
	else
		admin=false
		NoYieldchecker=false
		if instr(session("role"),",STORE_READER")=0 then
		response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
		end if
	end if
end if
%>