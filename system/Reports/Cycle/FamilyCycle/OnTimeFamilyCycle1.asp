<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'server.ScriptTimeout=2000%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
report_name=request.Form("report_name")
if request.Form("close_fromdate")<>"" then
close_from_time=request.Form("close_fromdate")&" "&request.Form("close_fromhour")&":"&request.Form("close_fromminute")
end if
if request.Form("close_todate")<>"" then
close_to_time=request.Form("close_todate")&" "&request.Form("close_tohour")&":"&request.Form("close_tominute")
end if

	set cmd=server.CreateObject("Adodb.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText="BAR_REPORT.RUN_WEEKLY_FAMILY_CYCLETIME"
	cmd.CommandType=4
	cmd.CommandTimeout=30
'	response.Write("REPORT_NAME := '"&replace(report_name,"'","^")&"';<br>OWNER_CODE := '"&session("code")&"';<br>OWNER_NAME := '"&session("user")&"';v_factory_id:=FA00000002;<br>START_TIME := '"&close_from_time&"';<br>END_TIME :='"&close_to_time&"';")
'	response.End()
	cmd.Parameters.Append cmd.CreateParameter("report_name", adVarChar, adParamInput, 200, replace(report_name,"'","^"))
	cmd.Parameters.Append cmd.CreateParameter("owner_code", adVarChar, adParamInput, 4, session("code"))
	cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
	cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 20, "FA00000002")
	cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, close_from_time)
	cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, close_to_time)
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