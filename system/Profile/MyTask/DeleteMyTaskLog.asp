<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
action="location.href='"&beforepath&"'"
rowid=request.QueryString("rowid")
SQL="delete from PROFILE_TASK_LOG where ROWID='"&rowid&"'"
rs.open SQL,conn,1,3
word="Successfully to delete Task Log.\n成功删除计划任务日志！"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->