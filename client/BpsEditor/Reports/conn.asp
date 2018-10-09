<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=10.6.100.51;uid=bde_client;pwd=bde_connect;database=MIS"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection") 
conn2.Open "driver={SQL Server};server=10.6.100.50;uid=sa;pwd=NXPisCool1;database=DWH"
%>