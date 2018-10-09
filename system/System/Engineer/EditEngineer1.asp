<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
id=request.Form("id")
username=request.Form("name")
userchinesename=request.Form("chinesename")
SQL="select * from USERS where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("USER_CODE")=trim(request.Form("code"))
rs("USER_NAME")=username
rs("USER_CHINESE_NAME")=userchinesename
rs("NT_ACCOUNT")=trim(request.Form("NT"))
rs("MANAGER")=trim(request.Form("manager"))
rs("FACTORY_ID")=replace(request.Form("factoryto")," ","")
rs("LANGUAGE")=request.Form("language")
rs("ROLES_ID")=replace(request.Form("toitem")," ","")
rs("EVENTS_ID")=replace(request.Form("event_id")," ","")
rs("EMAIL")=trim(request.Form("email"))
rs("APPROVAL_ROLE_ID")=request.Form("approval_role")
rs.update
word="Successfully edit Engineer."
action="location.href='"&beforepath&"'"
else
word="Engineer of "&username&" has not existed, please input again."
action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->