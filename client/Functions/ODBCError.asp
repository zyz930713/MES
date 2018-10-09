<%@ language="VBScript" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<head>
<META NAME="ROBOTS" CONTENT="NOINDEX">
<title>Page cannot be displayed</title>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=gb2312">
<META NAME="MS.LOCALE" CONTENT="ZH-CN">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="FFFFFF">
<table width="640" align="center" cellpadding="3" cellspacing="5">
  <tr>
    <td><p>System is under maintance. please wait.<br>
      it is estiamted that system will be recovered at <%= dateadd("n",30,now()) %> . <br>
        系统正在维护中，请等待。<br>
    预计你可以在 <%= dateadd("n",30,now()) %> 使用系统。</p>
    </td>
  </tr>
  <tr>
    <td><div align="center">
      <input type="button" name="Button" value="Home Page/首页" onClick="javascript:location.href='/Default.asp'">
    </div></td>
  </tr>
</table>
</body>
</html>