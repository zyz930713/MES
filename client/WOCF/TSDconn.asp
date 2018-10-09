

<%


Set connTSD = Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")
set rsA=server.CreateObject("adodb.recordset")
'connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.9,1433;User ID=bjsun;Password=Knowles01;Initial Catalog=TSD_Package" 

connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.153,1433;User ID=pvs;Password=pvs;Initial Catalog=TSD_Package" 



'Set conn = Server.CreateObject("ADODB.Connection")

'connstr = "driver={sql server};server=keb-vm101;uid=pvs;pwd=pvs;database=pvs"


'conn.Open connstr



%>

