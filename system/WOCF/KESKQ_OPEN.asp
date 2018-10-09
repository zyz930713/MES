<%
Set conn_web = Server.CreateObject("ADODB.Connection")
set rskq=server.createobject("adodb.recordset")
conn_web.Open "Provider=SQLOLEDB;Data Source=KES-VM-DB011;User ID=WebUser;Password=web2008it;Initial Catalog=HR50" 
%>