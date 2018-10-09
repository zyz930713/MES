<%
if connTicket.errors.count>0 then
connTicket.errors.clear
connTicket.rollbacktrans
end if
connTicket.committrans
%>