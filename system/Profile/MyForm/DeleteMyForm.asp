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
oracle_job_id=request.QueryString("job_id")

if oracle_job_id<>"" then
	set cmd=server.CreateObject("Adodb.Command") 
	cmd.ActiveConnection=conn 
	cmd.CommandText="DELETE_JOB"
	cmd.CommandType=4 
	cmd.Parameters.Append cmd.CreateParameter("n", adNumeric, adParamInput, 10, oracle_job_id)
	cmd.execute
	set cmd=nothing
	if err.number=0 then
	SQL="delete from PROFILE_FORM where NID='"&id&"' and USER_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	word="Successfully to delete System Form.\n成功删除系统表单和提醒任务！"
	else
	word="Fail to delete System Form.\n删除系统表单和提醒任务失败！"
	end if
else
SQL="delete from PROFILE_FORM where NID='"&id&"' and USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
word="Successfully to delete System Form.\n成功删除系统表单！"
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