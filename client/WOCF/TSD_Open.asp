<%
Set connTSD = Server.CreateObject("ADODB.Connection")
set connTSD001=Server.CreateObject("ADODB.Connection")
set rsTSD=server.createobject("adodb.recordset")

'connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.9,1433;User ID=bjsun;Password=Knowles01;Initial Catalog=TSD_Package" 
'connTSD.Open "Provider=SQLOLEDB;Data Source=KEB-TSD,1433;User ID=tsdweb;Password=Knowles01;Initial Catalog=KEB_TEST_RESULT_DB" 
if request("PRODUCT")="Explorer" then
connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.14,1433;User ID=tsd_reader;Password=tsd_reader;Initial Catalog=TSD_EXPLORER" 
else
connTSD.Open "Provider=SQLOLEDB;Data Source=10.6.89.153,1433;User ID=pvs;Password=pvs;Initial Catalog=TSD_Package" 
end if
set rsA=server.CreateObject("adodb.recordset")
%>