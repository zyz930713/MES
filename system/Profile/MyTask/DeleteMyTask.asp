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
id=request.QueryString("id")
on error resume next
job_id=request.QueryString("job_id")
if job_id<>"" then
set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="DELETE_JOB"
cmd.CommandType=4 
cmd.Parameters.Append cmd.CreateParameter("n", adNumeric, adParamInput, 10, job_id)
cmd.execute
set cmd=nothing
end if
if err.number=0 then
SQL="delete from PROFILE_TASK where NID='"&id&"'"
rs.open SQL,conn,1,3
word="Successfully to delete schedule.\n成功删除计划任务！"
else
word="Fail to delete schedule.\n删除计划任务失败！"
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