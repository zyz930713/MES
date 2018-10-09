<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Include/md5.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../styles/basic.css" />
<title>修改密码</title>
<style type="text/css">
form#Modify {
	width:280px;
	height:230px;
	border:1px solid #09f;
	margin:50px auto;
}
form#Modify h1 {
	height: 25px;
	line-height: 25px;
	background: #000066;
	font-size: 14px;
	text-align: center;
	color: #fff;
}
form#Modify label {
	display: block;
	padding-top: 8px;
	text-indent: 14px;
	font-weight: bold;
	color: #003366;
}
form#Modify label input.text {
	border:1px dashed #333;
	width:145px;
	height:19px;
	background:#fff;
}
#submit {
	width:64px;
	height:22px;
	margin:10px 0 0 54px;
	cursor:pointer;
}
</style>
</head>
<body bgcolor="#339966">
<%
HtmUrl = session("HtmUrl")
UserCode = session("UserCode")

if request("ModiPass") = "确认修改" then
	UserCode = Ucase(trim(request("UserCode")))
	password = md5(request("password"))
	passwordnew1 = md5(request("passwordnew1"))
	passwordnew2 = md5(request("passwordnew2"))
	
	if passwordnew1 <> passwordnew2 then
		response.write "<script>alert('重新输入的密码不同,请重新输入!');location.href='../../Functions/UserPass_Modify.asp'</script>"
		response.end()
	elseif sepasswordnew1 = "166c682bc41a473b" then
		response.write "<script>alert('请不要使用系统默认密码作为您的密码,请重新输入!');location.href='../../Functions/UserPass_Modify.asp'</script>"
		response.end()
	end if
	
	set rs = Server.CreateObject("adodb.recordset")
	SQL="select USER_CODE,USER_PASSWORD from USERS U where USER_CODE = '"&UserCode&"' and USER_PASSWORD = '"&password&"'"
	'response.write SQL
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<script>alert('原用户名或密码不正确,请重新输入!');</script>"
	else
		conn.execute("update USERS set USER_PASSWORD = '"&passwordnew1&"' WHERE USER_CODE = '"&UserCode&"'")
		'response.write "<script>alert('密码修改成功,下次请使用新密码登录!');</script>"
		' response.write "<script type='text/javascript'>"
		' response.write "window.opener=null;"
		' response.write "window.open('','_self');"
		' response.write "window.close();"
		' response.write "</script>"
		response.write "<script>alert('密码修改成功,下次请使用新密码登录!');location.href='"&HtmUrl&"'</script>"
		response.end()
		'response.redirect HtmUrl
		response.end()
	end if
end if
%>

<form name="Modify" id="Modify" method="post" action="UserPass_Modify.asp" >
	<h1>修改密码</h1>
	<label for="username">用　户　名：<input type="text" name="UserCode" value="<%=UserCode%>" id="username" class="text" readonly /></label>
	<label for="password">旧　密　码：<input type="password" name="password" id="password" class="text" /></label>
	<label for="password">新　密　码：<input type="password" name="passwordnew1" id="password" class="text" /></label>
	<label for="password">确认新密码：<input type="password" name="passwordnew2" id="password" class="text" /></label>
	<input type="submit" name="ModiPass" value="确认修改" id="submit" />
	<input type="button" value="放弃修改" onClick="window.location.reload('<%=HtmUrl%>');" id="submit" />
</form>
<br/>
</body>
</html>
