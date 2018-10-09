<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/MailHead.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
formid=request("formid")
fromsite=request.QueryString("fromsite")
'SQL="select PF.*,F.FORM_NAME,F.FORM_CHINESE_NAME,U.EMAIL from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID inner join USERS U on PF.USER_CODE=U.USER_CODE where PF.NID='"&formid&"' and PF.ACTOR_CODE like '%"&session("Code")&"%'"
SQL="select PF.*,F.FORM_NAME,F.FORM_CHINESE_NAME,U.EMAIL from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID inner join USERS U on PF.USER_CODE=U.USER_CODE where PF.NID='"&formid&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	applicant_user_code=rs("USER_CODE")
	applicant_mail=trim(rs("EMAIL"))
	form_name=rs("FORM_NAME")
	form_chinese_name=rs("FORM_CHINESE_NAME")
	if request("action")="1" then'accept to transact
		id=rs("FORM_ID")
		param1=trim(rs("PARAM1"))
		param2=trim(rs("PARAM2"))
		param3=trim(rs("PARAM3"))
		param4=trim(rs("PARAM4"))
		param5=trim(rs("PARAM5"))
		note=rs("NOTE")
		oracle_job_id=rs("ORACLE_JOB_ID")
		'run an oracle procedule
		set rsT=server.CreateObject("adodb.recordset")
			SQLT="select * from FORM where NID='"&id&"'"
			rsT.open SQLT,conn,1,3
			if not rsT.eof then
			packagename=rsT("PACKAGE")
			end if
			rsT.close
		set rsT=nothing
		if packagename<>"" then
			select case packagename
			case "JOB_QUANTITY_DECREASE"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="JOB_QUANTITY_DECREASE"
				cmd.CommandType=4
				'response.Write(param1&"@"&param2&"@"&session("code")&"@"&session("user")&"@"&param3)
				'response.End()
				'change_job_number in varchar2,new_quantity in varchar2,act_code in varchar2,act_name in nvarchar2,change_reason in nvarchar2,form_result out varchar2,stations_result out varchar2,canceled_sub_sheets out varchar2,error_info out nvarchar2
				cmd.Parameters.Append cmd.CreateParameter("form_id", adVarChar, adParamInput, 10, formid)
				cmd.Parameters.Append cmd.CreateParameter("change_job_number", adVarChar, adParamInput, 10, param1)
				cmd.Parameters.Append cmd.CreateParameter("new_quantity", adVarChar, adParamInput, 10, param2)
				cmd.Parameters.Append cmd.CreateParameter("act_code", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("act_name", adVarWChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("form_result", adVarChar, adParamOutput, 20)
				cmd.Parameters.Append cmd.CreateParameter("stations_result", adVarChar, adParamOutput, 50)
				cmd.Parameters.Append cmd.CreateParameter("canceled_sub_sheets", adVarChar, adParamOutput, 2000)
				cmd.Parameters.Append cmd.CreateParameter("error_info", adVarWChar, adParamOutput, 200)
				cmd.execute
				form_result=cmd("form_result")
				stations_result=cmd("stations_result")
				canceled_sub_sheets=cmd("canceled_sub_sheets")
				error_info=cmd("error_info")
				set cmd=nothing
				'response to client about run result.
				if err.number=0 then
					word="Scripts on server has executed! Result show as: <br>"
					cword="服务器脚步已执行！以下为运行结果：<br>"
					if form_result="true" then
						if stations_result<>"0" then
						word=word&"Info of stations is updated，total "&stations_result&" stations, please check again.<br>"
						cword=cword&"更新站的信息成功，共"&stations_result&"个站，请复查。<br>"
						else
						word=word&"Fail to update info of stations, please check again.<br>"
						cword=cword&"更新站的信息失败，请复查。<br>"
						end if
						if canceled_sub_sheets<>"" then
						word=word&"Following sub jobs are deleted: "&canceled_sub_sheets&"<br>"
						cword=cword&"以下细分工作单被删除："&canceled_sub_sheets&"<br>"
						else
						word=word&"No sub jobs are deleted.<br>"
						cword=cword&"没有细分工作单被删除。<br>"
						end if
						word=word&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/JobSheetsList.asp?jobnumber="&param1&"'>"&param1&"</a>"
						cword=cword&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/JobSheetsList.asp?jobnumber="&param1&"'>"&param1&"</a>"
					elseif form_result="false" then
						word=word&"No sub jobs are updated。"
						cword=cword&"没有细分工作单更新。"
					elseif form_result="No data" then
						word=word&"This job is not existed or startin to produce. Cannot be changed."
						cword=cword&"该工单不存在或没有开始生产，不能变更。"
					elseif form_result="Store Error" then
						word=word&"Stored quantity for job is exceed new quantity. Cannot be changed. Please contact retest or warehouse."
						cword=cword&"已经入库的数量大于新的工单数量，不能变更。请联系入库或仓库。"
					elseif form_result="Quantity Error" then
						word=word&"New quantity for job is exceed or equal to original quantity. Cannot be changed."
						cword=cword&"新的工单数量大于或等于原来数量，不能变更。"
					else
						word=word&"System error occurred and show as: <br>"&error_info
						cword=cword&"发生系统错误，以下为错误信息：<br>"&error_info
					end if
				else
					word="Fail to execute scripts on server! Please contact system administrator."
					cword="服务器脚步执行失败！请联系系统管理员。"
				end if			
			case "JOB_LINE_CHANGE"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="JOB_LINE_CHANGE"
				cmd.CommandType=4
				cmd.Parameters.Append cmd.CreateParameter("form_id", adVarChar, adParamInput, 20, formid)
				cmd.Parameters.Append cmd.CreateParameter("change_job_number", adVarChar, adParamInput, 10, param1)
				cmd.Parameters.Append cmd.CreateParameter("change_value", adVarChar, adParamInput, 10, param2)
				cmd.Parameters.Append cmd.CreateParameter("act_code", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("act_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("form_result", adVarChar, adParamOutput, 20)
				cmd.Parameters.Append cmd.CreateParameter("sheets_result", adVarChar, adParamOutput, 50)
				cmd.Parameters.Append cmd.CreateParameter("error_info", adVarChar, adParamOutput, 200)
				cmd.execute
				form_result=cmd("form_result")
				sheets_result=cmd("sheets_result")
				error_info=cmd("error_info")
				set cmd=nothing
				'response to client about run result.
				if err.number=0 then
					word="Scripts on server has executed! Result show as: <br>"
					cword="服务器脚步已执行！以下为运行结果：<br>"
					if form_result="1" then
						word=word&"Line info of sub jobs is updated, total "&sheets_result&" sheets. Please check again.<br>"
						cword=cword&"更新细分工作单的线别的信息成功，共"&sheets_result&"个细分工作单，请复查。<br>"
						word=word&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/JobSheetsList.asp?jobnumber="&param1&"'>"&param1&"</a>"
						cword=cword&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/JobSheetsList.asp?jobnumber="&param1&"'>"&param1&"</a>"
					else
						word=word&"System error occurred and show as: <br>"&error_info
						cword=cword&"发生系统错误，以下为错误信息：<br>"&error_info
					end if
				else
					word="Fail to execute scripts on server! Please contact system administrator."
					cword="服务器脚步执行失败！请联系系统管理员。"
				end if
			case "JOB_SCRAP_FORM"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="JOB_SCRAP_FORM"
				cmd.CommandType=4
				cmd.Parameters.Append cmd.CreateParameter("v_jobnumber", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("v_scrap_quantity", adVarChar, adParamInput, 500, param2)
				cmd.Parameters.Append cmd.CreateParameter("v_profile_form_id", adVarChar, adParamInput, 10, formid)
				cmd.Parameters.Append cmd.CreateParameter("v_code", adVarChar, adParamInput, 4, applicant_user_code)
				cmd.Parameters.Append cmd.CreateParameter("act_code", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("act_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("v_note", adVarChar, adParamInput, 200, note)
				cmd.Parameters.Append cmd.CreateParameter("v_reason", adVarChar, adParamInput, 50, param3)
				cmd.Parameters.Append cmd.CreateParameter("v_erp_param", adVarChar, adParamInput, 200, param4)
				cmd.Parameters.Append cmd.CreateParameter("v_oracle_job_id", adNumeric, adParamInput, 10, oracle_job_id)
				cmd.Parameters.Append cmd.CreateParameter("v_result", adNumeric, adParamOutput, 1)
				cmd.Parameters.Append cmd.CreateParameter("v_result_note", adVarChar, adParamOutput, 50)
				cmd.execute
				form_result=cmd("v_result")
				result_note=cmd("v_result_note")
				set cmd=nothing
				'response to client about run result.
				if err.number=0 then
					message="Scripts on server has executed! Result show as: <br>"
					cmessage="服务器脚步已执行！以下为运行结果：<br>"
					if form_result="1" then
						word=word&"Scrap "&param2&" for "&param1&" is saved. Please check detail.<br>"
						cword=cword&param1&"报废"&param2&"成功，请查看明细。<br>"
						word=word&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/ScrapRecords/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
						cword=cword&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/ScrapRecords/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
					else
						word=word&"System error occurred and show as: <br>"&result_note
						cword=cword&"发生系统错误，以下为错误信息：<br>"&result_note
					end if
				else
					message="Fail to execute scripts on server! Please contact system administrator."
					cmessage="服务器脚步执行失败！请联系系统管理员。"
				end if
			case "JOB_SCRAP_CHANGE_FORM"
				session("aerror")=formid
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="JOB_SCRAP_CHANGE_FORM"
				cmd.CommandType=4
				'response.Write(param2&"@"&param1&"@"&cint(param3)&"@"&formid&"@"&applicant_user_code&"@"&session("code")&"@"&session("user")&"@"&note&"@"&param4&"@"&oracle_job_id)
				'response.End()
				cmd.Parameters.Append cmd.CreateParameter("v_nid", adVarChar, adParamInput, 10, param2)
				cmd.Parameters.Append cmd.CreateParameter("v_jobnumber", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("v_scrap_quantity", adNumeric, adParamInput, 10, cint(param3))
				cmd.Parameters.Append cmd.CreateParameter("v_profile_form_id", adVarChar, adParamInput, 10, formid)
				cmd.Parameters.Append cmd.CreateParameter("v_code", adVarChar, adParamInput, 4, applicant_user_code)
				cmd.Parameters.Append cmd.CreateParameter("act_code", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("act_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("v_oracle_job_id", adNumeric, adParamInput, 10, oracle_job_id)
				cmd.Parameters.Append cmd.CreateParameter("v_reason", adVarWChar, adParamInput, 50, param4)
				cmd.Parameters.Append cmd.CreateParameter("v_erp_param", adVarChar, adParamInput, 200, param5)
				cmd.Parameters.Append cmd.CreateParameter("v_result", adNumeric, adParamOutput, 1)
				cmd.Parameters.Append cmd.CreateParameter("v_result_note", adVarChar, adParamOutput, 50)
				cmd.execute
				form_result=cmd("v_result")
				result_note=cmd("v_result_note")
				set cmd=nothing
				'response to client about run result.
				if err.number=0 then
					message="Scripts on server has executed! Result show as: <br>"
					cmessage="服务器脚步已执行！以下为运行结果：<br>"
					if form_result="1" then
						word=word&"Scrap "&param3&" for "&param1&" is saved. Please check detail.<br>"
						cword=cword&param1&"报废"&param3&"成功，请查看明细。<br>"
						word=word&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/ScrapRecords/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
						cword=cword&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/ScrapRecords/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
					else
						word=word&"System error occurred and show as: <br>"&result_note
						cword=cword&"发生系统错误，以下为错误信息：<br>"&result_note
					end if
				else
					message="Fail to execute scripts on server! Please contact system administrator."
					cmessage="服务器脚步执行失败！请联系系统管理员。"
				end if
			case "JOB_SCRAP_DELETE_FORM"
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="JOB_SCRAP_DELETE_FORM"
				cmd.CommandType=4
				cmd.Parameters.Append cmd.CreateParameter("v_nid", adVarChar, adParamInput, 10, param2)
				cmd.Parameters.Append cmd.CreateParameter("v_jobnumber", adVarChar, adParamInput, 20, param1)
				cmd.Parameters.Append cmd.CreateParameter("v_profile_form_id", adVarChar, adParamInput, 10, formid)
				cmd.Parameters.Append cmd.CreateParameter("v_code", adVarChar, adParamInput, 4, applicant_user_code)
				cmd.Parameters.Append cmd.CreateParameter("act_code", adVarChar, adParamInput, 4, session("code"))
				cmd.Parameters.Append cmd.CreateParameter("act_name", adVarChar, adParamInput, 50, session("user"))
				cmd.Parameters.Append cmd.CreateParameter("v_note", adVarChar, adParamInput, 200, note)
				cmd.Parameters.Append cmd.CreateParameter("v_reason", adVarChar, adParamInput, 50, param4)
				cmd.Parameters.Append cmd.CreateParameter("v_oracle_job_id", adNumeric, adParamInput, 10, oracle_job_id)
				cmd.Parameters.Append cmd.CreateParameter("v_result", adNumeric, adParamOutput, 1)
				cmd.Parameters.Append cmd.CreateParameter("v_result_note", adVarChar, adParamOutput, 50)
				cmd.execute
				form_result=cmd("v_result")
				result_note=cmd("v_result_note")
				set cmd=nothing
				'response to client about run result.
				if err.number=0 then
					message="Scripts on server has executed! Result show as: <br>"
					cmessage="服务器脚步已执行！以下为运行结果：<br>"
					if form_result="1" then
						word=word&"Scrap "&param3&" for "&param1&" is saved. Please check detail.<br>"
						cword=cword&param1&"报废"&param3&"成功，请查看明细。<br>"
						word=word&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/Scrap/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
						cword=cword&"<a href='http://"&request.ServerVariables("HTTP_HOST")&"/Job/Scrap/ScrapRecords.asp?jobnumber="&param1&"'>"&param1&"</a>"
					else
						word=word&"System error occurred and show as: <br>"&result_note
						cword=cword&"发生系统错误，以下为错误信息：<br>"&result_note
					end if
				else
					message="Fail to execute scripts on server! Please contact system administrator."
					cmessage="服务器脚步执行失败！请联系系统管理员。"
				end if
			end select
		end if
		set rs1=server.CreateObject("adodb.recordset")
		SQL1="update PROFILE_FORM set FORM_STATUS=4 where NID='"&formid&"'"
		rs1.open SQL1,conn,1,3
		set rs1=nothing
	else'reject to transact
		set rs1=server.CreateObject("adodb.recordset")
		SQL1="update PROFILE_FORM set FORM_STATUS=6,REJECT_REASON='"&request.Form("rejectreason")&"' where NID='"&formid&"'"

		rs1.open SQL1,conn,1,3
		set rs1=nothing
		if applicant_mail<>"" then
			sendEmail application("MailSender"),applicant_mail,"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />Your form <"&form_name&"> is rejected to be transacted by "&session("user")&".你申请的表单<"&form_chinese_name&">被"&session("user")&"拒绝执行。",mailhead&"<p>Hi, Group Leader<p>Your form <"&form_name&"> is rejected to be transacted by "&session("user")&".你申请的表单<"&form_chinese_name&">被"&session("user")&"拒绝执行。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/MyFormSee.asp?formid="&formid&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
			word="Rejected change form. System sent email notification to "&applicant_mail&"."
			cword="拒绝变更成功。已发送电子邮件通知给"&applicant_mail&"。"
		else
			word="Rejected change form. But no email address for transactor, please notify him/her by telephone."
			cword="拒绝变更成功。但没有申请人的邮件地址，请电话通知。"
		end if
	end if
else
	word="No authority to access or it is not existed or is transaced."
	cword="无权访问或不存在或已执行。"
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Form Transaction/表单处理</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="700" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000066" bordercolordark="#FFFFFF">
  <tr>
	<td height="20" class="t-c-greenCopy"><%if session("language")="0" then%><%=message%><%else%><%=cmessage%><%end if%></td>
  </tr>
  <tr>
    <td height="20"><%if session("language")="0" then%>
      <%=word%>
      <%else%>
      <%=cword%>
    <%end if%></td>
  </tr>
  <tr>
	<td height="20"><div align="center">
	  <%if fromsite="board" then%>
	  <input type="button" name="Button" value="<%if session("language")="0" then%>Return<%else%>返回<%end if%>" onClick="javascript:location.href='/Profile/Myform/FormBoard.asp'">
	  <%else%>
	  <input type="button" name="Button" value="<%if session("language")="0" then%>Close<%else%>关闭<%end if%>" onClick="javascript:window.close()">
	  <%end if%>
	</div></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->