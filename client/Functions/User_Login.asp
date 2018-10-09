<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Include/md5.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户登录</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<style type="text/css">
.InputText{
	width:145px;
	height:19px;
}
#LoginForm{
	height:260px;
}
</style>
</head>
<body style="margin:20px;background-color:#339966;" onLoad="javascript:document.all.UserCode.focus();">
<%
HtmUrl = session("HtmUrl")

UserCode = UCASE(trim(request("UserCode")))
PassWord = MD5(request("PassWord"))
SubmitLogin = request("SubmitLogin")



if SubmitLogin = "登录" then
	SQL="select U.*,F.FACTORY_NAME from USERS U left join FACTORY F on U.FACTORY_ID = F.NID where U.USER_CODE = '"&UserCode&"' and USER_PASSWORD = '"&PassWord&"'"
	'response.write SQL
	
	rs.open SQL,conn,1,1
	if not rs.eof then
		if rs("STATUS")<>"1" then
			response.Redirect("/Error.asp?error=Error message: your account is disabled.您的账号已被停用。")
		end if
		session("UserCode")=rs("USER_CODE")
		session("user")= rs("USER_NAME")
		session("factory")=rs("FACTORY_ID")
		session("factory_name")=rs("FACTORY_NAME")
		session("role")=","&rs("ROLES_ID")
		session("email")=rs("EMAIL")
		session("language")=rs("LANGUAGE")
		session("series_filter")=rs("SERIES_FILTER")
		session("yieldexclusion_filter")=rs("YIELDEXCLUSION_FILTER")
		if PassWord = "166c682bc41a473b" then
			response.write "<script>alert('首次登录,请修改密码!');location.href='../../Functions/UserPass_Modify.asp'</script>"
			response.end()
		else
			response.redirect HtmUrl
			response.end()
		end if
	else
		response.write "<script>alert('用户名或密码错误,请重新输入!');location.href='../../Functions/user_login.asp'</script>"
		response.end()
	end if
	rs.close
end if
%>
<div id="LoginForm">
<center>
<form name="form1" method="post" action="user_login.asp" >
<table width="320" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">用户登录</div>
		</td>
	</tr>	
	<tr>
		<td align="center">
			<br/>
			<table width="240" border="0" align="center" cellpadding="0" cellspacing="0" >
			<tr>
				<td><br/>
				用户名：<input type="text" name="UserCode" id= "UserCode"  class="InputText" onKeyDown="if(event.keyCode==13){event.keyCode=9;}" />
				</td>
			</tr>
			<tr>
				<td><br/>
				密　码：<input type="password" name="PassWord" class="InputText" />
				</td>
			</tr>
			<tr>
				<td><br/>
				<input type="submit" name="SubmitLogin" value="登录" />　
				<input type="reset" value="重填" />
				</td>
			</tr>
			</table>
			<br/>
		</td>
	</tr>
</table>
</form>
</center>
</div>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->