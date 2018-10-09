<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'server.ScriptTimeout=2000%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
report_name=request.Form("report_name")
part_number=request.Form("part_number")
job_number=request.Form("job_number")
line_name=request.Form("line_name")
job_status=request.Form("job_status")
if request.Form("start_fromdate")<>"" then
start_from_time=request.Form("start_fromdate")&" "&request.Form("start_fromhour")&":"&request.Form("start_fromminute")
end if
if request.Form("start_todate")<>"" then
start_to_time=request.Form("start_todate")&" "&request.Form("start_tohour")&":"&request.Form("start_tominute")
end if
if request.Form("close_fromdate")<>"" then
close_from_time=request.Form("close_fromdate")&" "&request.Form("close_fromhour")&":"&request.Form("close_fromminute")
end if
if request.Form("close_todate")<>"" then
close_to_time=request.Form("close_todate")&" "&request.Form("close_tohour")&":"&request.Form("close_tominute")
end if

	set cmd=server.CreateObject("Adodb.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText="BAR_REPORT.RUN_ONTIME_JOB_INFORMATION"
	cmd.CommandType=4
	cmd.CommandTimeout=30   
	'response.Write("REPORT_NAME := '"&replace(report_name,"'","^")&"';<br>OWNER_CODE := '"&session("code")&"';<br>OWNER_NAME := '"&session("user")&"';<br>PART_NUMBER := '"&ucase(part_number)&"';<br>JOB_NUMBER := '"&ucase(job_number)&"';<br>LINE_NAME := '"&ucase(line_name)&"';<br>JOB_STATUS := '"&job_status&"';<br>START_START_TIME := '"&start_from_time&"';<br>START_END_TIME := '"&start_to_time&"';<br>CLOSE_START_TIME := '"&close_from_time&"';<br>CLOSE_END_TIME :='"&close_to_time&"';")
'	response.End()
	cmd.Parameters.Append cmd.CreateParameter("report_name", adVarChar, adParamInput, 200, replace(report_name,"'","^"))
	cmd.Parameters.Append cmd.CreateParameter("owner_code", adVarChar, adParamInput, 4, session("code"))
	cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
	cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 50, ucase(part_number))
	cmd.Parameters.Append cmd.CreateParameter("job_number", adVarChar, adParamInput, 50, ucase(job_number))
	cmd.Parameters.Append cmd.CreateParameter("line_name", adVarChar, adParamInput, 50, ucase(line_name))
	cmd.Parameters.Append cmd.CreateParameter("job_status", adVarChar, adParamInput, 1, job_status)
	cmd.Parameters.Append cmd.CreateParameter("start_start_time", adVarChar, adParamInput, 20, start_from_time)
	cmd.Parameters.Append cmd.CreateParameter("start_end_time", adVarChar, adParamInput, 20, start_to_time)
	cmd.Parameters.Append cmd.CreateParameter("close_start_time", adVarChar, adParamInput, 20, close_from_time)
	cmd.Parameters.Append cmd.CreateParameter("close_end_time", adVarChar, adParamInput, 20, close_to_time)
	cmd.Parameters.Append cmd.CreateParameter("error_string", adVarChar, adParamOutput,500)
	cmd.execute
	if err.number=0 then
	word="You will recieve report in few minutes!"
	else
	error_string=cmd("error_string")
	word="Fail to generate report！"&error_string
	end if
	set cmd=nothing
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->