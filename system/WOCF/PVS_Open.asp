<%
'Set connPVS = Server.CreateObject("ADODB.Connection")
'set rsPVS=server.createobject("adodb.recordset")
'connPVS.Open "Provider=SQLOLEDB;Data Source=10.6.89.153,1433;User ID=pvs;Password=pvs;Initial Catalog=pvs"



Set conn = Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
conn.Open "Provider=SQLOLEDB;Data Source=10.6.89.153,1433;User ID=pvs;Password=pvs;Initial Catalog=pvs" 

 
%>