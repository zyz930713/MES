<%
if session("role")<>"" then
	if instr(session("role"),",ACTION_ADMINISTRATOR")<>0 then
	admin=true
	else
		admin=false
		if instr(session("role"),",ACTION_READER")=0 then
		response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
		end if
	end if
else
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if

ActionPurpose=Array("Other","Machine Code","Material Part Number","Material Lot Number","Material Quantity","Job Quantity","Rework Quantity","2D Code","Defect 2D Code")
%>