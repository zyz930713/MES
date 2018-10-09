<%
set connTicket = Server.CreateObject("ADODB.Connection")
set rsTicket=server.createobject("adodb.recordset")
'Dev
'ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome1;User ID=BAR_PROD_KEB;Data Source=KES1DEV"

'Prod
ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome3;User ID=BAR_PROD_KEB;Data Source=KEBDB"


connTicket.Open ConnString
%>