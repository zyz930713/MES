<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=10.6.100.51;uid=bde_client;pwd=bde_connect;database=MIS"

Dim ConnOra
set ConnOra = Server.CreateObject("ADODB.Connection")
ConnOra.Open "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDB"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection") 
conn2.Open "driver={SQL Server};server=10.6.100.50;uid=sa;pwd=NXPisCool1;database=DWH"
%>