<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
taskname=trim(request.Form("taskname"))
SQL="delete from TASK where NID='"&id&"'"
rs.open SQL,conn,1,3
word="Successfully delete Task."
action="location.href='"&beforepath&"'"
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