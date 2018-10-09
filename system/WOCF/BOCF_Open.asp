
<%
set conn = Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
'Dev
'ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome1;User ID=BAR_WEB_KEB;Data Source=KES1DEV"
'Prod
ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDB"

conn.Open ConnString
%>