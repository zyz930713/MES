<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=10.6.100.44;uid=zhaoran;pwd=Philips1;database=KEB_ReportData"
%>