<%
Dim ConnTSD01
Set ConnTSD01 = Server.CreateObject("ADODB.Connection")
ConnTSD01.Open "Provider=SQLOLEDB;Data Source=10.6.89.14,1433;User ID=TSD_EXPLORER;Password=TSD_EXPLORER;Initial Catalog=TSD_EXPLORER" 
%>