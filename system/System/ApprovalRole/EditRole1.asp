<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
SQL="select * from SYSTEM_APPROVAL_ROLE where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
rs("ROLE_NAME")=rolename
rs("ROLE_CHINESE_NAME")=trim(request.Form("chinesename"))
rs("DESCRIPTION")=trim(request.Form("description"))
rs.update
rs.close
word="Successfully Approval edit Role. Update "&update_count&"; Cancel "&cancel_count
action="location.href='"&beforepath&"'"
else
word="Approval Role of "&rolename&" has existed, please input again."
action="history.back()"
rs.close
end if
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