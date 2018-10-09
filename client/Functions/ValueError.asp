<%@ language="VBScript" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<head>
<META NAME="ROBOTS" CONTENT="NOINDEX">
<title>Page cannot be displayed</title>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=gb2312">
<META NAME="MS.LOCALE" CONTENT="ZH-CN">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=10;
adCount=0;
function showPopup()
{
adCount+=1
	if(adCount<timePopup)
	{
	setTimeout("showPopup()",1000);
	document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
	else
	{
	closePopup()
	}
}
function closePopup()
{
location.href="/Sisonic/Station_Close.asp"
}
</script>
</head>

<body  onLoad="showPopup();">
<table width="640" align="center" cellpadding="3" cellspacing="5">
  <tr>
    <td><div align="center"><span class="strongred">Window will close window in 10 seconds.&nbsp;<span id="countinsert"></span></span></div></td>
  </tr>
  <tr>
    <td><p>Values of following actions are invid or unchecked, please return to try again: 下列数值输入错误或未校验，请返回重试：<br>
      <%= request.QueryString("erroractions") %></p>    </td>
  </tr>
  <tr>
    <td><div align="center">
      <input type="button" name="Button" value="Home Page/首页" onClick="javascript:location.href='/Default.asp'">
    </div></td>
  </tr>
</table>
</body>
</html>