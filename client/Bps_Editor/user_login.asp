﻿<!--#include file="include/md5.asp"-->
<!--#include file= "Conn_open.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../styles/basic.css" />
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<title>BPS editor login</title>
<style type="text/css">
form#login {
	width: 250px;
	height: 170px;
	border: 1px solid #40BF80;
	margin: 150px auto;
}
form#login h1 {
	height: 25px;
	line-height: 25px;
	background: #000066;
	font-size: 16px;
	text-align: center;
	color: #fff;
}
form#login label {
	display: block;
	padding-top: 8px;
	text-indent: 14px;
	color: #003366;
	font-weight: bold;
}
form#login label input.text {
	border: 1px dashed #40BF80;
	width: 140px;
	height: 19px;
	background: #fff;
}
form#login input.submit {
	width: 60px;
	height: 19px;
	border: 1px dashed #40BF80;
	background: #fff;
	margin: 10px 0 0 100px;
	cursor: pointer;
}
</style>
</head>
<body bgcolor="#339966">
<%
if request("send") = "Login" then
	seusername = lcase(request("seusername"))
	sepassword = lcase(request("sepassword"))

	sepassword = md5(sepassword)
	'response.write sepassword
	'response.end
	
	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT UID,Eusername,Epassword,Egroup,Etype from report_userlist WHERE Eusername = '"&seusername&"' AND Epassword = '"&sepassword&"'"
	rs.open sql,conn,1,1
	if rs.eof then
		response.write "<script>alert('Incorrect user name or password!');</script>"
	else
		session("ShiftEditUser") = seusername
		session("ShiftEditRole") = rs("Egroup")
		session("ShiftEditType") = rs("Etype")
		Etype = session("ShiftEditType")
		Egroup = session("ShiftEditRole")
		if Etype = "OEM" then response.redirect "InputOem.asp"
		
		if Etype = "HV" then
			if Egroup = "0" then
				response.redirect "InputHV.asp"
			else
				response.redirect "HvMachine.asp"
			end if
		end if
		
		if Etype = "HVPER" then response.redirect "InputPeriphery.asp"
		if Etype = "TARGET" then response.redirect "SetHv.asp"
	end if

end if
%>
<form method="post" action="user_login.asp" name="login" id="login">
	
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td colspan="2"><h1>Login in</h1></td>
    </tr>
    <tr>
      <td><label for="username">Username：</label></td>
      <td><input type="text" name="seusername" id="username" class="text" /></td>
    </tr>
    <tr>
      <td><label for="password">Password：</label></td>
      <td><input type="password" name="sepassword" id="password" class="text" /></td>
    </tr>
    <tr>
      <td colspan="2"><input type="submit" name="send" value="Login" class="submit" /></td>
    </tr>
  </table>
</form>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
</body>
</html>
