<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'response.Redirect("/KES1/Station.asp")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System 站点</title>
<script language="javascript">
function pageload()
{
	window.opener=null;
	KES1=window.open('/KES1/Station1.asp','KES1','fullscreen=yes',true);
	KES1.focus();
	window.close();
}
</script>
</head>

<body onload="pageload()">
</body>
</html>