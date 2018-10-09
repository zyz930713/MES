<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=2000%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
profile_task_id=request.QueryString("profile_task_id")
factory_id=request.QueryString("factory_id")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")

SQL="select PT.NID,PT.TASK_ID,PT.USER_CODE,T.TASK_NAME,U.USER_NAME from PROFILE_TASK PT inner join TASK T on PT.TASK_ID=T.NID left join USERS U on PT.USER_CODE=U.USER_CODE where PT.NID='"&profile_task_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then

	if factory_id="FA00000002" then
		set cmd=server.CreateObject("Adodb.Command")
		cmd.ActiveConnection=conn
		cmd.CommandText="RUN_DAILY_LINEYIELD"
		cmd.CommandType=4
'		response.Write(profile_task_id&"#"&rs("TASK_ID")&"#"&rs("TASK_NAME")&"#"&rs("USER_CODE")&"#"&rs("USER_NAME")&"#"&factory_id&"#"&from_time&"#"&to_time)
'		response.End()
		cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, profile_task_id)
		cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, rs("TASK_ID"))
		cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, rs("TASK_NAME"))
		cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, rs("USER_CODE"))
		cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, rs("USER_NAME"))
		cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 20, factory_id)
		cmd.Parameters.Append cmd.CreateParameter("start_hour", adVarChar, adParamInput, 20, from_time)
		cmd.Parameters.Append cmd.CreateParameter("end_hour", adVarChar, adParamInput, 20, to_time)
		cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, null)
		cmd.execute
		set cmd=nothing
	else
	set cmd=server.CreateObject("Adodb.Command")
		cmd.ActiveConnection=conn
		cmd.CommandText="RUN_DAILY_LINEYIELD_BYSERIES"
		cmd.CommandType=4 
		cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, profile_task_id)
		cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, rs("TASK_ID"))
		cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, rs("TASK_NAME"))
		cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, rs("USER_CODE"))
		cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, rs("USER_NAME"))
		cmd.Parameters.Append cmd.CreateParameter("v_factory_id", adVarChar, adParamInput, 20, factory_id)
		cmd.Parameters.Append cmd.CreateParameter("start_hour", adVarChar, adParamInput, 20, from_time)
		cmd.Parameters.Append cmd.CreateParameter("end_hour", adVarChar, adParamInput, 20, to_time)
		cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, null)
		cmd.execute
		set cmd=nothing
	end if
end if
rs.close
word="Regeneration is OK!"
action="window.close()"
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