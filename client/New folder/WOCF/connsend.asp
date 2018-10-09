<%

'Set connSend = Server.CreateObject("ADODB.Connection")
'set rs=server.createobject("adodb.recordset")
'connSend.Open "Provider=SQLOLEDB;Data Source=KEB-dt056;User ID=sa;Password=Admin888;Initial Catalog=NPISend" 




Set connSend = Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
connSend.Open "Provider=SQLOLEDB;Data Source=KEB-2DB;User ID=IT;Password=Knowles02;Initial Catalog=NPISend" 



%>

