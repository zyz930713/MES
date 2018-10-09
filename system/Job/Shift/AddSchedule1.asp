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

line_id=request.Form("line_id")
line_name=request.Form("line_name")
shift_type=request.Form("shift_type")
runday=request.Form("day")
runhour=request.Form("hour")
runminute=request.Form("minute")
runtime=runday&" "&runhour&":"&runminute

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="CREATE_SHIFT_JOB"
cmd.CommandType=4 
cmd.Parameters.Append cmd.CreateParameter("v_type", adVarChar, adParamInput, 4, shift_type)
cmd.Parameters.Append cmd.CreateParameter("v_line_id", adVarChar, adParamInput, 10, line_id)
cmd.Parameters.Append cmd.CreateParameter("v_line_name", adVarChar, adParamInput, 10, line_name)
cmd.Parameters.Append cmd.CreateParameter("v_person", adVarChar, adParamInput, 4, session("code") )
cmd.Parameters.Append cmd.CreateParameter("v_runtime", adVarChar, adParamInput, 20, runtime)
cmd.execute
'response.Write(cmd("o_test1"))
set cmd=nothing
if err.number=0 then
word="Successfully to create schedule.\n成功创建计划任务！"
else
word="Fail to create schedule.\n创建计划任务失败！"
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