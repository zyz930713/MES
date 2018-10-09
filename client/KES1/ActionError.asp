<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Sisonic Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20"><div align="center" class="strongred">Alert Message</div></td>
  </tr>
  <tr>
    <td height="20" class="t-t-DarkBlue">
	<%select case request("alerttype") 
	case "lock"
	alerttip="Following code, machine or material is lock, operation fails, please contact engineer.<br>下列工号或机器或原材料被锁定，本次操作失败，请联系工程师。"
	case "part"
	alerttip="Following part of material is not in BOM, please try again.<br>下列原材料型号不在BOM中，本次操作失败，请联系工程师。"
	case "conjuctive"
	alerttip="Conjuctive job has not closed, please try again.<br>附属工单还没有结束，请联系工程师。"
	end select%><%=alerttip%></td>
  </tr>
  
  <tr>
    <td height="20"><% =request("alertmessage") %></td>
  </tr>
  <tr>
    <td height="20"><div align="center">
      <input type="button" name="Button" value="Home Page/首页" onClick="javascript:location.href='/KES1/Station1.asp'">
    </div></td>
  </tr>
</table>
</body>
</html
><%session.Abandon()%>