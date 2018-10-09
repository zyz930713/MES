<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=keb-vm100;uid=sa;pwd=Knowles01;database=KEB_PVS"
%>