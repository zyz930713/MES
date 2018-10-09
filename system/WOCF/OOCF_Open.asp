<%
Set OraSession=CreateObject("OracleInProcServer.XOraSession")
Set OraDatabase=OraSession.DbOpenDatabase("KES1Bar","bar_web/barcode",0)
%>