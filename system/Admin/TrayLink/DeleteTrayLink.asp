<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TrayLink/TrayLinkCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
partnumber=request.QueryString("partnumber")
stationid=request.QueryString("stationid")
traytype=request.QueryString("traytype")
SQL="delete from part_tray where part_number='"&partnumber&"' and station_id='"&stationid&"' and tray_type='"&traytype&"'"
rs.open SQL,conn,3,3
SystemLog "Admin - TrayLink","Delete TrayLink of "&stationid&" - "&traytype&" ("&partnumber&")"
word="TrayLink is deleted!"
action="history.back()"
%>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
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
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->