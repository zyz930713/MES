<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="currenttime(thisnow)" background="/Images/un.gif">
<table border="0" cellpadding="0" cellspacing="0">
	<tr> 
	  <td valign="top"><img src="/Images/knowles-logo.jpg" width="1020" height="82"></td>	 
	</tr>
	<tr>
	  <td nowrap><div align="right"> </div><div align="right">Beijing Time: <span id=currenttimeinsert></span>&nbsp;</div></td>
	</tr>
</table>
</body>
</html>
<script language="javascript">
initime=new Date('<%
	z=month(date)&"-"&day(date)&"-"&year(date)&space(1)&formatdatetime(now,3)
	%><%=z%>');
thisnow=initime.getTime()
function currenttime(thistime)
{
displaytime=new Date(thistime);
document.all.currenttimeinsert.innerText=displaytime.getFullYear()+"-"+(displaytime.getMonth()+1)+"-"+displaytime.getDate()+" "+displaytime.getHours()+":"+displaytime.getMinutes()+":"+displaytime.getSeconds();
thistime=thistime+1000;
window.setTimeout("currenttime("+thistime.toString()+")",1000); 
}
</script>
