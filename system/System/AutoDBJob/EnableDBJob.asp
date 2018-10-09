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

on error resume next
id=request.QueryString("id")

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="ENABLE_SYSTEM_JOB"
cmd.CommandType=4 
cmd.Parameters.Append cmd.CreateParameter("n", adNumeric, adParamInput, 10, id)
cmd.execute
set cmd=nothing
if err.number=0 then
word="Successfully to enable schedule.\n成功生效计划任务！"
else
word="Fail to enable schedule.\n生效计划任务失败！"
end if
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