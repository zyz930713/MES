<%
if conn.errors.count>0 then
conn.errors.clear
conn.rollbacktrans
end if
conn.committrans
%>