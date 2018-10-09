<%
Set connTSD = Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")
'connTSD.Open "Provider=SQLOLEDB;Data Source=KEB-TSD,1433;User ID=pvs;Password=pvs;Initial Catalog=pvs" 
'connTSD.Open "Provider=SQLOLEDB;Data Source=keb-tsd,1433;User ID=tsdweb;Password=Knowles01;Initial 
connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.9,1433;User ID=bjsun;Password=Knowles01;Initial Catalog=TSD_Package" 

'Catalog=KEB_TEST_RESULT_DB" 

%>