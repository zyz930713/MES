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
isnew=false
report_name=request.Form("report_name")
from_time=request.Form("fromdate")
to_time=request.Form("todate")
factory=request.Form("factory")
family=request.Form("family")
report_type=request.Form("report_type")
this_station=request.Form("this_station")
this_station_name=request.Form("this_station_name")
this_action=request.Form("this_action")
this_action_name=request.Form("this_action_name")
refer_station=request.Form("refer_station")
refer_station_name=request.Form("refer_station_name")
refer_action=request.Form("refer_action")
refer_action_name=request.Form("refer_action_name")
recievers=replace(request.Form("toitem")," ","")

set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="JOB_ACTION_TRACKING"
cmd.CommandType=4
'response.Write("THIS_FACTORY_ID := '"&factory&"';<br>  THIS_FAMILY_ID := '"&family&"';<br>  THIS_ACTION_ID := '"&this_action&"';<br>  THIS_ACTION_NAME := '"&this_action_name&"';<br>THIS_STATION_ID := '"&this_station&"';<br>  THIS_STATION_NAME := '"&this_station_name&"';<br>  FROM_TIME := '"&from_time&"';<br>  TO_TIME := '"&to_time&"';<br>  TRACKING_TYPE := '"&report_type&"';<br>  REFER_ACTION_ID := '"&refer_action&"';<br>  REFER_ACTION_NAME := '"&refer_action_name&"';<br>  REFER_STATION_ID := '"&refer_station&"';<br>  REFER_STATION_NAME := '"&refer_station_name&"';<br>  REPORT_NAME := '"&report_name&"';<br>  SENDER := '"&session("user")&"';<br>  RECIEVERS := '"&recievers&"';")
'response.end
cmd.Parameters.Append cmd.CreateParameter("this_factory_id", adVarChar, adParamInput, 10, factory)
cmd.Parameters.Append cmd.CreateParameter("this_family_id", adVarChar, adParamInput, 10, family)
cmd.Parameters.Append cmd.CreateParameter("this_action_id", adVarChar, adParamInput, 10, this_action)
cmd.Parameters.Append cmd.CreateParameter("this_action_name", adVarChar, adParamInput, 100, this_action_name)
cmd.Parameters.Append cmd.CreateParameter("this_station_id", adVarChar, adParamInput, 10, this_station)
cmd.Parameters.Append cmd.CreateParameter("this_station_name", adVarChar, adParamInput, 100, this_station_name)
cmd.Parameters.Append cmd.CreateParameter("from_time", adVarChar, adParamInput, 10, from_time)
cmd.Parameters.Append cmd.CreateParameter("to_time", adVarChar, adParamInput, 10, to_time)
cmd.Parameters.Append cmd.CreateParameter("tracking_type", adVarChar, adParamInput, 30, report_type)
cmd.Parameters.Append cmd.CreateParameter("refer_action_id", adVarChar, adParamInput, 10, refer_action)
cmd.Parameters.Append cmd.CreateParameter("refer_action_name", adVarChar, adParamInput, 100, refer_action_name)
cmd.Parameters.Append cmd.CreateParameter("refer_station_id", adVarChar, adParamInput, 10, refer_station)
cmd.Parameters.Append cmd.CreateParameter("refer_station_name", adVarChar, adParamInput, 100, refer_station_name)
cmd.Parameters.Append cmd.CreateParameter("report_name", adVarChar, adParamInput, 100, report_name)
cmd.Parameters.Append cmd.CreateParameter("sender", adVarChar, adParamInput, 20,  session("user"))
cmd.Parameters.Append cmd.CreateParameter("recievers", adVarChar, adParamInput, 2000, recievers)
cmd.Parameters.Append cmd.CreateParameter("error_message", adVarChar, adParamOutput, 500)
cmd.execute
set cmd=nothing
if err.number=0 then
word="\n立即运行任务成功！"
else
word="\n立即运行任务失败！请联系系统管理员。"
end if
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