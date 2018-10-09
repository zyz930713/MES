<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
isnew=false
id=request.Form("task")
taskname=trim(request.Form("taskname"))
taskname=replace(taskname,"'","^")
recievers=replace(request.Form("toitem")," ","")
atonce=request.Form("atonce")
period=request.Form("period")
schedule_type=request.Form("happenitem")
if request.Form("hour_interval")<>"" then
hour_interval=cdbl(request.Form("hour_interval"))
else
hour_interval=""
end if
if request.Form("week") then
week_day=request.Form("week")
else
week_day=""
end if
if request.Form("param1")<>"" then
param1=replace(request.Form("param1")," ","")
	if request.Form("param_type1")="time" then
	param1=replace(param1,",",":")
	elseif request.Form("param_type1")="period" then
	param1="P|"&replace(param1,",","|")
	end if
else
param1=""
end if
if request.Form("param2")<>"" then
param2=replace(request.Form("param2")," ","")
	if request.Form("param_type2")="time" then
	param2=replace(param2,",",":")
	elseif request.Form("param_type2")="period" then
	param2="P|"&replace(param2,",","|")
	end if
else
param2=""
end if
if request.Form("param3")<>"" then
param3=replace(request.Form("param3")," ","")
	if request.Form("param_type3")="time" then
	param3=replace(param3,",",":")
	elseif request.Form("param_type3")="period" then
	param3="P|"&replace(param3,",","|")
	end if
else
param3=""
end if
if request.Form("param4")<>"" then
param4=replace(request.Form("param4")," ","")
	if request.Form("param_type4")="time" then
	param4=replace(param4,",",":")
	elseif request.Form("param_type4")="period" then
	param4="P|"&replace(param4,",","|")
	end if
else
param4=""
end if
runtime=request.Form("fromdate")&" "&request.Form("happentime1")&":"&request.Form("happentime2")
NID="PT"&NID_SEQ("PROFILETASK")

SQL="select * from PROFILE_TASK where USER_CODE='"&session("code")&"' and TASK_ID='"&id&"' and TASK_NAME='"&taskname&"'"
rs.open SQL,conn,1,3
if rs.eof then
isnew=true
else
isnew=false
end if
rs.close

if isnew=true then
	'create a job
	set rsT=server.CreateObject("adodb.recordset")
		SQLT="select * from TASK where NID='"&id&"'"
		rsT.open SQLT,conn,1,3
		if not rsT.eof then
		packagename=rsT("PACKAGE")
		end if
		rsT.close
	set rsT=nothing
	if packagename<>"" then
		if atonce="1" then
			select case packagename
			case "UNSTORE_JOB_ALERT"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="UNSTORE_JOB_ALERT"
				cmd.CommandType=4 
				cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, null)
				cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, id)
				cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, taskname)
				cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("days", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("job_scope", adVarChar, adParamInput, 20, param2)
				cmd.Parameters.Append cmd.CreateParameter("factory_scope", adVarChar, adParamInput, 20, param3)
				cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, recievers)
				cmd.execute
				set cmd=nothing
			case "UNFINISHED_JOB_ALERT"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="UNFINISHED_JOB_ALERT"
				cmd.CommandType=4 
				cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, null)
				cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, id)
				cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, taskname)
				cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("days", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("line_scope", adVarChar, adParamInput, 20, param2)
				cmd.Parameters.Append cmd.CreateParameter("factory_scope", adVarChar, adParamInput, 50, param3)
				cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, recievers)
				cmd.execute
				set cmd=nothing
			case "PRODUCTS_WIP_ALERT"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="PRODUCTS_WIP_ALERT"
				cmd.CommandType=4
				'response.Write(id&"@"&taskname&"@"&session("code")&"@"&session("user")&"@"&param1&"@"&param2&"@"&recievers)
				'response.End()
				cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, null)
				cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, id)
				cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, taskname)
				cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("line_scope", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("factory_scope", adVarChar, adParamInput, 50, param2)
				cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, recievers)
				cmd.execute
				set cmd=nothing
			end select
			if err.number=0 then
			word="\n立即运行任务成功！"
			else
			word="\n立即运行任务失败！请联系系统管理员。"
			end if
		end if
		if period="1" then
			set cmd=server.CreateObject("Adodb.Command") 
			cmd.ActiveConnection=conn 
			cmd.CommandText="CREATE_TASK_JOB"
			cmd.CommandType=4 
			'response.Write(NID&"@"&taskname&"@"&session("code")&"@"&session("user")&"@"&packagename&"@"&schedule_type&"@"&week_day&"@"&param1&"@"&param2&"@"&recievers&"@"&runtime)
			'response.End()
			cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, NID)
			cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, id)
			cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 200, taskname)
			cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, session("code"))
			cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, session("user"))
			cmd.Parameters.Append cmd.CreateParameter("package_name", adVarChar, adParamInput, 200, packagename)
			cmd.Parameters.Append cmd.CreateParameter("schedule_type", adVarChar, adParamInput, 20, schedule_type)
			cmd.Parameters.Append cmd.CreateParameter("weekday", adVarChar, adParamInput, 20, week_day)
			cmd.Parameters.Append cmd.CreateParameter("hour_interval", adDouble, adParamInput, 20, hour_interval)
			cmd.Parameters.Append cmd.CreateParameter("param1", adVarChar, adParamInput, 20, param1)
			cmd.Parameters.Append cmd.CreateParameter("param2", adVarChar, adParamInput, 20, param2)
			cmd.Parameters.Append cmd.CreateParameter("param3", adVarChar, adParamInput, 20, param3)
			cmd.Parameters.Append cmd.CreateParameter("param4", adVarChar, adParamInput, 20, param4)
			cmd.Parameters.Append cmd.CreateParameter("param5", adVarChar, adParamInput, 20, param5)
			cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, recievers)
			cmd.Parameters.Append cmd.CreateParameter("runtime", adVarChar, adParamInput, 30, runtime)
			cmd.Parameters.Append cmd.CreateParameter("job_id", adInteger, adParamOutput)
			cmd.execute
			oracle_job_id=cmd("job_id")
			set cmd=nothing
			if err.number=0 then
			word="\n创建计划任务成功！"
			else
			word="\n创建计划任务失败！请联系系统管理员。"
			end if
		end if
	end if
	
	SQL="select * from PROFILE_TASK where USER_CODE='"&session("code")&"' and TASK_ID='"&id&"' and TASK_NAME='"&taskname&"'"
	rs.open SQL,conn,1,3
	rs.addnew
	rs("NID")=NID
	rs("USER_CODE")=session("code")
	rs("TASK_ID")=id
	rs("TASK_NAME")=taskname
	rs("RECIEVERS")=recievers
		if param1<>"" then
		rs("PARAM1")=param1
		end if
		if param2<>"" then
		rs("PARAM2")=param2
		end if
		if param3<>"" then
		rs("PARAM3")=param3
		end if
		if param4<>"" then
		rs("PARAM4")=param4
		end if
		if period="1" then
		rs("IS_SCHEDULE")=period
		rs("SCHEDULE_TYPE")=schedule_type
			if request.Form("happenitem")="hour" then
			rs("HOUR_INTERVAL")=hour_interval
			elseif request.Form("happenitem")="week" then
			rs("WEEK_DAY")=week_day
			end if
		rs("START_TIME")=request.Form("fromdate")
		rs("HAPPEN_TIME")=request.Form("happentime1")&":"&request.Form("happentime2")
		rs("ORACLE_JOB_ID")=oracle_job_id
		end if
	rs.update
	rs.close
	word="Successfully save new Task.成功保存新任务。"&word
	action="location.href='MyTask.asp'"
else
	word="Task has existed."
	action="location.href='"&beforepath&"'"
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