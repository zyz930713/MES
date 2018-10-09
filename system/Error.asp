<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<html>
<head>
<title>Error Message</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<table border="0" align="center" cellpadding="0" cellspacing="0">
  <%if request.QueryString("error")<>"" then%>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <%end if%>
  <tr > 
    <td height="20"><div align="center">Welcome to <%=application("SystemName")%></div></td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td height="20"><span class="nred"><%=request.QueryString("error")%></span></td>
  </tr>
</table>
</body>
</html>
