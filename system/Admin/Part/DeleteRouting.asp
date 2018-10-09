<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
partnumber=request.QueryString("partnumber")

SQL="update Routing set IS_DELETE=1,LASTUPDATE_PERSON='"&session("code")&"',LASTUPDATE_TIME='"&date()&"' where NID='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from ROUTING_ACTION_DETAIL where Routing_Id='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from ROUTING_Defect_DETAIL where Routing_Id='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from ACTION where Routing_Id='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from Defectcode where Routing_Id='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from PART where Mother_Routing_ID='"&id&"'"
rs.open SQL,conn,3,3
SQL="delete from Station where Routing_Id='"&id&"'"
rs.open SQL,conn,3,3


'SystemLog "Admin - Part","Disable Part of "&partnumber&" ("&id&")"
word="Routing of "&partnumber&" is deleted!"
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
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