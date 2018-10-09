<%
Set connTSD = Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")

connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.9,1433;User ID=bjsun;Password=Knowles01;Initial Catalog=TSD_PackageMarigold" 
set rsA=server.CreateObject("adodb.recordset")
%>