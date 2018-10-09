<%
if instr(session("role"),",OBI_ADMINISTRATOR")<>0 then
OBI=true
else
OBI=false
end if
%>