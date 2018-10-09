<%
Set connIA = Server.CreateObject("ADODB.Connection")
set rsIA=server.createobject("adodb.recordset")

'connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.9,1433;User ID=bjsun;Password=Knowles01;Initial Catalog=TSD_Package" 
'connTSD.Open "Provider=SQLOLEDB;Data Source=KEB-TSD,1433;User ID=tsdweb;Password=Knowles01;Initial Catalog=KEB_TEST_RESULT_DB" 
'connTSD.Open "Provider=SQLOLEDB;Data Source=KEB-VM101,1433;User ID=pvs;Password=pvs;Initial Catalog=KEB_TEST_RESULT_DB" 
connIA.Open "Provider=SQLOLEDB;Data Source=10.6.89.14,1433;User ID=TSD_EXPLORER;Password=TSD_EXPLORER;Initial Catalog=TSD_EXPLORER" 
set rsA=server.CreateObject("adodb.recordset")
%>