<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then
	call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
else
	seusername = session("ShiftEditUser")
end if
%>
<!--#include file= "include/function.asp"-->
<!--#include file="include/md5.asp"-->
<!--#include file= "Conn_open.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../styles/basic.css" />
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<title>OEM Shift editer login</title>
<style type="text/css">
form#login {
	width:280px;
	height:230px;
	border:1px solid #09f;
	margin:150px auto;
}
form#login h1 {
	height: 25px;
	line-height: 25px;
	background: #000066;
	font-size: 14px;
	text-align: center;
	color: #fff;
}
form#login label {
	display: block;
	padding-top: 8px;
	text-indent: 14px;
	font-weight: bold;
	color: #003366;
}
form#login label input.text {
	border:1px dashed #333;
	width:160px;
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
if request("send") = "确认修改" then
	seusername = request.form("seusername")
	sepassword = md5(request.form("sepassword"))
	sepasswordnew1 = md5(request.form("sepasswordnew1"))
	sepasswordnew2 = md5(request.form("sepasswordnew2"))
	'response.write sepassword
	'response.end
	if sepasswordnew1 <> sepasswordnew2 then call sussLoctionHref("新密码与确认新密码必须相同,请重新输入!","modify_pass.asp")
	
	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT UID,Eusername,Epassword,Egroup,Etype from report_userlist WHERE Eusername = '"&seusername&"' AND Epassword = '"&sepassword&"'"
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<script>alert('原用户名或密码不正确,请重新输入!');</script>"
	else
		conn.execute("UPDATE report_userlist SET Epassword = '"&sepasswordnew1&"' WHERE Eusername = '"&seusername&"'")
		call sussLoctionHref("密码修改完毕,请下次使用新密码登陆!","default.asp")
	end if

end if
%>

<form method="post" action="modify_pass.asp" name="login" id="login">
	<h1>修改密码</h1>
  <label for="username">用　户　名：<input type="text" name="seusername" value="<%=seusername%>" id="username" class="text" readonly /></label>
  <label for="password">旧　密　码：<input type="password" name="sepassword" id="password" class="text" /></label>
  <label for="password">新　密　码：<input type="password" name="sepasswordnew1" id="password" class="text" /></label>
  <label for="password">确认新密码：<input type="password" name="sepasswordnew2" id="password" class="text" /></label>
	<input type="button" value="放弃修改" onClick="window.location.reload('default.asp');" id="submit" /><input type="submit" name="send" value="确认修改" id="submit" />
</form>
</body>
</html>
