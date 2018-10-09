<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
partnumber=request.QueryString("partnumber")
SQL="update PART set STATUS=0 where NID='"&id&"'"
rs.open SQL,conn,3,3
SystemLog "Admin - Part","Disable Part of "&partnumber&" ("&id&")"
word="Part of "&partnumber&" is disabled!"
%>
<script language="javascript">
alert("<%=word%>");
location.href='<%=beforepath%>';
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->