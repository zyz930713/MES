<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
username=trim(request.Form("name"))
userchinesename=trim(request.Form("chinesename"))
NT=trim(request.Form("NT"))
SQL="select * from USERS where NT_ACCOUNT='"&NT&"'"
rs.open SQL,conn,3,3
if rs.eof then
rs.addnew
rs("NID")="US"&NID_SEQ("USER")
rs("USER_CODE")=trim(request.Form("code"))
rs("USER_NAME")=username
rs("USER_CHINESE_NAME")=userchinesename
rs("NT_ACCOUNT")=NT
rs("FACTORY_ID")=replace(request.Form("factoryto")," ","")
if request.Form("language")<>"" then
rs("LANGUAGE")=request.Form("language")
end if
rs("ROLES_ID")=replace(request.Form("toitem")," ","")
rs("EVENTS_ID")=replace(request.Form("event_id")," ","")
rs("EMAIL")=trim(request.Form("email"))
rs("APPROVAL_ROLE_ID")=request.Form("approval_role")
rs("USER_PASSWORD")="166c682bc41a473b"
rs.update
word="Successfully save New Engineer."
action="location.href='"&beforepath&"'"
else
word="Engineer of "&username&" has existed, please input again."
action="history.back()"
end if
rs.close
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<script language="JavaScript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->