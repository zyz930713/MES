<%
if session("role")<>"" then
	if instr(session("role"),",FINANCE_SERIESGROUP_ADMINISTRATOR")<>0 then
	admin=true
	else
		admin=false
		if instr(session("role"),",FINANCE_SERIESGROUP_READER")=0 then
		response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>