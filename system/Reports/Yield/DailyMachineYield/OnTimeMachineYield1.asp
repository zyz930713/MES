<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=2000%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
report_name=request.Form("report_name")
factory_id=request.Form("factory_id")
station=request.Form("station")
if station<>"" then
station=replace(station," ","")
end if
family=request.Form("family")
series=request.Form("series")
model_name=request.Form("model_name")
line_name=request.Form("line_name")
machine_name=request.Form("machine_name")
from_time=request.Form("fromdate")&" "&request.Form("fromhour")&":"&request.Form("fromminute")
to_time=request.Form("todate")&" "&request.Form("tohour")&":"&request.Form("tominute")
by_day=request.Form("by_day")

	set cmd=server.CreateObject("Adodb.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText="BAR_REPORT.RUN_ONTIME_MACHINE_YIELD"
	cmd.CommandType=4
'	response.Write(replace(report_name,"'","^")&"#"&session("code")&"#"&session("user")&"#"&factory_id&"#"&station&"#"&family&"#"&series&"#"&model_name&"#"&line_name&"#"&machine_name&"#"&from_time&"#"&to_time)
'	response.End()
	cmd.Parameters.Append cmd.CreateParameter("report_name", adVarChar, adParamInput, 200, replace(report_name,"'","^"))
	cmd.Parameters.Append cmd.CreateParameter("owner_code", adVarChar, adParamInput, 4, session("code"))
	cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
	cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 20, factory_id)
	cmd.Parameters.Append cmd.CreateParameter("v_station", adVarChar, adParamInput, 20, station)
	cmd.Parameters.Append cmd.CreateParameter("v_family", adVarChar, adParamInput, 20, family)
	cmd.Parameters.Append cmd.CreateParameter("v_series", adVarChar, adParamInput, 20, series)
	cmd.Parameters.Append cmd.CreateParameter("v_model_name", adVarChar, adParamInput, 50, model_name)
	cmd.Parameters.Append cmd.CreateParameter("v_line_name", adVarChar, adParamInput, 20, line_name)
	cmd.Parameters.Append cmd.CreateParameter("v_machine_name", adVarChar, adParamInput, 20, machine_name)
	cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, from_time)
	cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, to_time)
	cmd.Parameters.Append cmd.CreateParameter("by_day", adVarChar, adParamInput, 1, by_day)
	cmd.execute
	set cmd=nothing

word="You will recieve report in few minutes!"
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