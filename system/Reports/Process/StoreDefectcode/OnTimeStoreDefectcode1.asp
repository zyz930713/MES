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
if request.Form("store_fromdate")<>"" then
store_from_time=request.Form("store_fromdate")&" "&request.Form("store_fromhour")&":"&request.Form("store_fromminute")
end if
if request.Form("store_todate")<>"" then
store_to_time=request.Form("store_todate")&" "&request.Form("store_tohour")&":"&request.Form("store_tominute")
end if

	set cmd=server.CreateObject("Adodb.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText="BAR_REPORT.RUN_ONTIME_STORE_DEFECTCODE"
	cmd.CommandType=4
	cmd.CommandTimeout=30   
'	response.Write("REPORT_NAME := '"&replace(report_name,"'","^")&"';<br>OWNER_CODE := '"&session("code")&"';<br>OWNER_NAME := '"&session("user")&"';<br>PART_NUMBER := '"&ucase(part_number)&"';<br>JOB_NUMBER := '"&ucase(job_number)&"';<br>LINE_NAME := '"&ucase(line_name)&"';<br>STORE_START_TIME := '"&store_from_time&"';<br>STORE_END_TIME := '"&store_to_time&"';")
'	response.End()
	cmd.Parameters.Append cmd.CreateParameter("report_name", adVarChar, adParamInput, 200, replace(report_name,"'","^"))
	cmd.Parameters.Append cmd.CreateParameter("owner_code", adVarChar, adParamInput, 4, session("code"))
	cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
	cmd.Parameters.Append cmd.CreateParameter("part_number", adVarChar, adParamInput, 50, ucase(part_number))
	cmd.Parameters.Append cmd.CreateParameter("job_number", adVarChar, adParamInput, 50, ucase(job_number))
	cmd.Parameters.Append cmd.CreateParameter("line_name", adVarChar, adParamInput, 50, ucase(line_name))
	cmd.Parameters.Append cmd.CreateParameter("store_start_time", adVarChar, adParamInput, 20, store_from_time)
	cmd.Parameters.Append cmd.CreateParameter("store_end_time", adVarChar, adParamInput, 20, store_to_time)
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