<%
	if session("ShiftEditUser") = "" then
		response.redirect "user_login.asp"
	else
		Etype = session("ShiftEditType")
		if Etype = "OEM" THEN
			response.redirect "InputOem.asp"
		elseif Etype = "HV" then
			response.redirect "InputHv.asp"
		elseif Etype = "HVPER" then
			response.redirect "InputPeriphery.asp"
		end if
	end if
%>