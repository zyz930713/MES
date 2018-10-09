<%
if session("role")<>"" then
	if instr(session("role"),",AMOUNT_ADMINISTRATOR")<>0 then
	amount_admin=true
	else
		amount_admin=false
		if instr(session("role"),",AMOUNT_READER")=0 then
		amount_reader=true
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>