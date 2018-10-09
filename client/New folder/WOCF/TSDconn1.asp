

<%


Set connTSD = Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")
set rsA=server.CreateObject("adodb.recordset")
connTSD.Open "Provider=SQLOLEDB;Data Source=KEB-TSD,1433;User ID=tsdweb;Password=Knowles01;Initial Catalog=KEB_TEST_RESULT_DB" 



'Set conn = Server.CreateObject("ADODB.Connection")

'connstr = "driver={sql server};server=keb-vm101;uid=pvs;pwd=pvs;database=pvs"


'conn.Open connstr



%>

