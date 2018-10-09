<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
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
seriesname=request.QueryString("seriesname")
SQL="update SUBSERIES set STATUS=1 where NID='"&id&"'"
rs.open SQL,conn,3,3
SystemLog "Admin - Sub Series","Enable Sub Series of "&seriesname&" ("&id&")"
word="Sub Series of "&seriesname&" is enabled!"
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