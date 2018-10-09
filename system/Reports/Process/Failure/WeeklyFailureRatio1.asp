<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=200%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
factory_id=request("factory_id")
year_number=request("year_number")
week_number=request("week_number")
if request("report_type")="" then
from_time=request.Form("fromdate")&" "&request.Form("fromhour")&":"&request.Form("fromminute")
to_time=request.Form("todate")&" "&request.Form("tohour")&":"&request.Form("tominute")
else
from_time=request.Form("from_time")
to_time=request.Form("to_time")
end if

	set cmd=server.CreateObject("Adodb.Command")
	cmd.ActiveConnection=conn
	cmd.CommandText="BAR_REPORT.RUN_WEEKLY_FAILURE_RATIO"
	cmd.CommandType=4
'	response.Write(year_number&"#"&week_number&"#"&session("code")&"#"&session("user")&"#"&factory_id&"#"&from_time&"#"&to_time)
'	response.End()
	cmd.Parameters.Append cmd.CreateParameter("v_year_number", adVarChar, adParamInput, 4, year_number)
	cmd.Parameters.Append cmd.CreateParameter("v_week_number", adVarChar, adParamInput, 4, week_number)
	cmd.Parameters.Append cmd.CreateParameter("owner_code", adVarChar, adParamInput, 50, session("code"))
	cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 20, factory_id)
	cmd.Parameters.Append cmd.CreateParameter("start_time", adVarChar, adParamInput, 20, from_time)
	cmd.Parameters.Append cmd.CreateParameter("end_time", adVarChar, adParamInput, 20, to_time)
	cmd.execute
	set cmd=nothing

word="Report is generated successfully!"
action="location.href='WeeklyStationYieldList.asp?factory_id="&factory_id&"&year_number="&year_index&"&week_number="&week_number
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