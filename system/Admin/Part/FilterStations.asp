<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Functions/GetStation.asp" -->
<%
key=request.QueryString("key")
where=replace(session("filterwhere"),"where"," and ")
options=getStation(true,"OPTION",""," where S.STATION_NAME like '%"&key&"%' "&where,"","","")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
parent.document.all.insert_stations.innerHTML="<select name='fromitem2' size='10' multiple id='fromitem2'><%=options%></select>"
</script>
</head>

<body>
</body>
</html>
