<%
set conn = Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDBT"
conn.Open ConnString
%>