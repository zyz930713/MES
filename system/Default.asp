<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<html>
<head>
<title><%=application("SystemName")%> - Home Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<%if session("code")="" then
'取得用户NT帐号
varname=lcase(request.ServerVariables("remote_USER"))

'使用winnt协议取得用户完全帐号名称
	SQL="select U.*,F.FACTORY_NAME from USERS U left join FACTORY F on U.FACTORY_ID=F.NID where lower(U.NT_ACCOUNT)='"&lcase(varname)&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		if rs("STATUS")<>"1" then
			response.Redirect("/Error.asp?error=Error message: your account is disabled.您的账号已被停用。")
		end if
		session("code")=rs("USER_CODE")
		session("user")= rs("USER_NAME")
		session("factory")=rs("FACTORY_ID")
		session("factory_name")=rs("FACTORY_NAME")
		session("role")=","&rs("ROLES_ID")
		session("email")=rs("EMAIL")
		session("language")=rs("LANGUAGE")
		session("series_filter")=rs("SERIES_FILTER")
		session("yieldexclusion_filter")=rs("YIELDEXCLUSION_FILTER")		
	else
		response.Redirect("/Error.asp?error=Error message: you are not authorized to access this page.您没有权限访问此页面。NT_ACCOUNT="&varname)
	end if
	rs.close
end if%>
<frameset rows="110,*" cols="*" framespacing="1" frameborder="NO" border="1">
  <frame src="/Components/Nav.asp" name="top" frameborder="no" scrolling="NO" noresize id="top">
  <frameset rows="*" cols="161,*" framespacing="1" frameborder="no" border="1">
    <frame src="/Components/Left.asp" name="left" frameborder="no" scrolling="auto" id="left">
    <frame src="/Main.asp" name="main" frameborder="no" scrolling="auto" id="main">
  </frameset>
</frameset>
<noframes>
<body>

</body></noframes>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
