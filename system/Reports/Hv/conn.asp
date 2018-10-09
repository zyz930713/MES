<%
Dim ConnSql
Set ConnSql = Server.CreateObject("ADODB.Connection")
ConnSql.Open "driver={SQL Server};server=10.6.100.44;uid=zhaoran;pwd=Philips1;database=KEB_ReportData"

Dim ConnOra
set ConnOra = Server.CreateObject("ADODB.Connection")
ConnOra.Open "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDB"
%>