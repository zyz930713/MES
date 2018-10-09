<%
if session("role")<>"" then
	if instr(session("role"),",SERIESGROUP_ADMINISTRATOR")<>0 then
	admin=true
	capacity=false
	elseif instr(session("role"),",SERIESGROUP_CAPACITY")<>0 then
	capacity=true
	admin=false
	else
		admin=false
		capacity=false
		if instr(session("role"),",SERIESGROUP_READER")=0 then
		response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>