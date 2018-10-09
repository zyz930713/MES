<%
set conn = Server.CreateObject("ADODB.Connection")
ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDBT"
conn.Open ConnString

Dim ConnSql
Set ConnSql = Server.CreateObject("ADODB.Connection")
ConnSql.Open "driver={SQL Server};server=10.6.100.44;uid=zhaoran;pwd=Philips1;database=KEB_ReportData" 
%>