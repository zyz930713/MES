<%
set rsF=server.CreateObject("adodb.recordset")
Set Upload = Server.CreateObject("Persits.Upload")
Upload.LogonUser "", "ASPEncrypt", "Dickens;'2006"
Upload.CodePage = 936
Upload.Save
%>
