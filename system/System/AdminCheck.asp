<%
if instr(session("role"),",SYSTEM_ADMINISTRATOR")<>0 then
sysadmin=true
else
sysadmin=false
end if
%>