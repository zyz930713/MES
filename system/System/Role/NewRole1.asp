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
rolename=trim(request.Form("name"))
SQL="select * from ENGINEER_ROLE where ROLE_NAME='"&rolename&"'"
rs.open SQL,conn,3,3
if rs.eof then
rs.addnew
rs("NID")="RO"&NID_SEQ("ROLE")
rs("ROLE_NAME")=rolename
rs("ROLE_CHINESE_NAME")=trim(request.Form("chinesename"))
rs("DESCRIPTION")=trim(request.Form("description"))
rs("APPLICANT")=trim(request.Form("applicant"))
if request.Form("apply")="1" then
rs("APPLY_FROM_WEB")=1
else
rs("APPLY_FROM_WEB")=0
end if
rs.update
word="Successfully save New Role."
action="location.href='"&beforepath&"'"
else
word="Role of "&rolename&" has existed, please input again."
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