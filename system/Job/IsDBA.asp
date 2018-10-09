<%
if instr(session("role"),",DBA")<>0 then
DBA=true
else
DBA=false
end if
%>