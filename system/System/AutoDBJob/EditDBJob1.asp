<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
action="location.href='"&beforepath&"'"

id=request.Form("id")
jobitem=request.Form("jobitem")
period=request.Form("period")
runday=request.Form("day")
runhour=request.Form("hour")
runminute=request.Form("minute")
runtime=runday&" "&runhour&":"&runminute

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="EDIT_SYSTEM_JOB"
cmd.CommandType=4 
cmd.Parameters.Append cmd.CreateParameter("v_job_id", adNumeric, adParamInput, 20, id)
cmd.Parameters.Append cmd.CreateParameter("v_job_item", adVarChar, adParamInput, 100, jobitem)
cmd.Parameters.Append cmd.CreateParameter("v_runtime", adVarChar, adParamInput, 20, runtime)
cmd.Parameters.Append cmd.CreateParameter("v_period", adVarChar, adParamInput, 100, period)
cmd.Parameters.Append cmd.CreateParameter("v_person", adVarChar, adParamInput, 10, session("code"))
cmd.execute
set cmd=nothing
if err.number=0 then
word="Successfully to edit DB job.\n成功修改计划任务！"
action="location.href='"&beforepath&"'"
else
word="Fail to edit DB job.\n修改计划任务失败！"
action="history.back()"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
alert("<%=word%>")
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->