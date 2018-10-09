<%
if session("role")<>"" then
	if instr(session("role"),",STORE_ADMINISTRATOR")<>0 then
	admin=true
	retestchecker=false
		if instr(session("role"),",STORE_RETESTCHECKER")<>0 then
		retestchecker=true
		end if
	else
		if instr(session("role"),",STORE_RETESTCHECKER")<>0 then
		admin=false
		retestchecker=true
		else
			admin=false
			retestchecker=false
			if instr(session("role"),",STORE_READER")=0 then
			response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
			end if
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>