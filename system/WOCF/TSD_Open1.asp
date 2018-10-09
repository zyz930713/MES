<%
Set connTSD = Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")

connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.153,1433;User ID=pvs;Password=pvs;Initial Catalog=TSD_Package" 

%>