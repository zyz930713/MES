<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/MailHead.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/Functions/GetFormUser.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/GetScrapCost.asp" -->
<!--#include virtual="/Functions/GetUserReportTo.asp" -->
<!--#include virtual="/Functions/GetUserAgent.asp" -->
<!--#include virtual="/Functions/GetUserInfo.asp" -->
<%
path=request("path")
query=request("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
factory=request.QueryString("factory")
id=request("thisform")
param1=trim(request("param1"))
param2=trim(request("param2"))
param3=trim(request("param3"))
param4=trim(request("param4"))
param5=trim(request("param5"))
fromsite=trim(request("fromsite"))
note=request("note")
group_leader=trim(request("group_leader"))
group_leader_mail=trim(request("group_leader_mail"))
code=trim(request("code")&"")

'get info about form
SQL="select * from FORM where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
packagename=rs("PACKAGE")
form_name=rs("FORM_NAME")
form_chinese_name=rs("FORM_CHINESE_NAME")
form_type=rs("FORM_TYPE")
approve1="3"
'approve2=rs("APPROVE2")
alert_interval=cint(rs("ALERT_TIME"))
alert_person=rs("ALERT_PERSON")
end if
rs.close
'get info about job
'get approve flow info
if form_type="1" then '1=need to be approved; 0=don't need to be approved.
	if approve1<>"" then
		select case approve1
		case "3" 'supervisor of line
		getFormUser "FORMUSER","where USER_CODE=(select L.SUPERVISOR from LINE L inner join JOB_MASTER J on L.LINE_NAME=J.LINE_NAME where J.JOB_NUMBER='"&param1&"')",user_code,user_name,email
		approve_code=user_code
		last_approve_code=approve_code
		approve_name=user_name
		approve_email=email
		approve_time=" "
		approve_status="0"
		current_approve_code=user_code
		current_approve_name=user_name
		current_approve_email=approve_email
		'current_approve_email="dickens.xu@knowles.com"
		end select
		
		if approve2<>"" then
		select case approve2
		case "5" 'user's director manager
		getFormUser "FORMUSER","where USER_CODE=(select MANAGER from USERS where USER_CODE='"&approve_code&"')",user_code,user_name,email
		approve_code=approve_code&","&user_code
		last_approve_code=approve_code
		approve_name=approve_name&","&user_name
		approve_email=approve_email&","&email
		approve_invertal=interval
		approve_time=approve_time&","&" "
		approve_status=approve_status&","&"0"
		end select
		end if
	else
	response.Redirect("http://"&application("ClientServer")&"/Scrap/Scrap1.asp?factory="&factory&"&errorstring=该"&jobnumber&"工单没有设定线别主管，请联系系统管理员！")
	end if
	'get alert settings
	if alert_interval<>0 and alert_person<>"" then
	alert_time=dateadd("n",alert_interval,now())
		select case alert_person
		case "5"
		getFormUser "FORMUSER","where USER_CODE=(select MANAGER from USERS where USER_CODE='"&approve_code&"')",user_code,user_name,email
		alert_code=user_code
		alert_name=user_name
		alert_email=email
		end select
	end if
end if

if id="FO00000006" then '报废
	
		arrQty=split(param2,",")
		totalQty=0
		j=0
		
		for i=0 to ubound(arrQty)
			j=trim(arrQty(i)&"")
			if j="" or (not isnumeric(j)) then
				j=0
			end if
			totalQty=cdbl(totalQty) + cdbl(j)
			
		next
		
		item_cost=GetItemCostByJob(param1)
		
		'ScrapCost=cdbl(item_cost) * cdbl(totalQty)
		'for OBI Test
		ScrapCost=0
		
		dim approve(7)
		dim amount(7)
		set rst=server.CreateObject("adodb.recordset")

		strUserCode=code
		
		strUserEName=GetUserInfo(strUserCode,"ENGLISHNAME")
		strUserCName=GetUserInfo(strUserCode,"NAME")
		
		strPRE=GetUserReportTo(strUserCode,"code")
		strPREName=GetUserReportTo(strUserCode,"ename")
		isPRE=false
		
		
		SQL="select * from USERS where USER_CODE='" & strPRE & "' and APPROVAL_ROLE_ID='AR00000001'"
		rs.open SQL,conn,1,3
		if rs.eof then
			strPRE=GetUserReportTo(strPRE,"code")
		else
			isPRE=true				
		end if
		rs.close
		
		'strApproveList=strPRE
		strApproveList=""
	
		if isPRE=false then
			strPRE=GetUserReportTo(strUserCode,"code")
		end if		
		
		SQL="select * from FORM where NID='FO00000006'"
		
		rs.open SQL,conn,1,3
		
		if not rs.eof then
			if trim(rs("approve1") & "")<>"" then
				approve(0)=trim(rs("approve1") & "")
				amount(0)=trim(rs("amount1") & "")
				if trim(rs("amount1") & "")="" then
					amount(0)=0
				end if
			end if
			if trim(rs("approve2") & "")<>"" then
				approve(1)=trim(rs("approve2") & "")
				amount(1)=trim(rs("amount2") & "")
				if trim(rs("amount2") & "")="" then
					amount(1)=0
				end if		
			end if
			
			if trim(rs("approve3") & "")<>"" then
				approve(2)=trim(rs("approve3") & "")
				amount(2)=trim(rs("amount3") & "")
				if trim(rs("amount3") & "")="" then
					amount(2)=0
				end if		
			end if
			if trim(rs("approve4") & "")<>"" then
				approve(3)=trim(rs("approve4") & "")
				amount(3)=trim(rs("amount4") & "")
				if trim(rs("amount4") & "")="" then
					amount(3)=0
				end if		
			end if
			if trim(rs("approve5") & "")<>"" then
				approve(4)=trim(rs("approve5") & "")
				amount(4)=trim(rs("amount5") & "")
				if trim(rs("amount5") & "")="" then
					amount(4)=0
				end if		
			end if
			if trim(rs("approve6") & "")<>"" then
				approve(5)=trim(rs("approve6") & "")
				amount(5)=trim(rs("amount6") & "")
				if trim(rs("amount6") & "")="" then
					amount(5)=0
				end if		
			end if
			if trim(rs("approve7") & "")<>"" then
				approve(6)=trim(rs("approve7") & "")
				amount(6)=trim(rs("amount7") & "")
				if trim(rs("amount7") & "")="" then
					amount(6)=0
				end if		
			end if
					
		end if
		rs.close
		
		if GetUserAgent(strPRE)<>"" then
		strPRE=GetUserAgent(strPRE)
		end if
		strApproveList=strPRE
		
		
		strDirectManager=strPRE
		
		strOtherManager=""
		
		iGrade=1
		
		for i=1 to 6
			'response.Write(ScrapCost & "  VS " & amount(i) & "  VS " & approve(i) & "<br>")
			if trim(approve(i)&"")<>"" and cdbl(ScrapCost)>=cdbl(amount(i)) then
				
				'strApproveList=strApproveList & "," & approve(i)
				
				strDirectManager=GetCurrentManager(strDirectManager,"code")
				
				SQL="select * from USERS where USER_CODE='" & strDirectManager & "' and APPROVAL_ROLE_ID='" & trim(approve(i)&"") & "'"
				rs.open SQL,conn,1,3
				if rs.eof then
					
					SQL="select * from USERS where APPROVAL_ROLE_ID='" & trim(approve(i)&"") & "' order by APPROVAL_ROLE_SEQ"
					rst.open SQL,conn,1,3
					if not rst.eof then
						strOtherManager=trim(rst("USER_CODE")&"")
						if trim(GetUserAgent(strOtherManager)&"")<>"" then
							strOtherManager=GetUserAgent(strOtherManager)
						end if
					end if
					rst.close
					strApproveList=strApproveList & "," & strOtherManager
				else
					strApproveList=strApproveList & "," & strDirectManager	
				end if
				rs.close
				iGrade=iGrade+1
			end if	
		next
		
		'response.Write(strApproveList & "---" & iGrade)
		
		if iGrade=1 then
			LastApprover=strPRE
		else
			LastApprover=mid(strApproveList,instrrev(strApproveList,",")+1)
		end if


		'response.Write(item_cost & " -- " & param1 & " -- " & totalQty & "---" & strApproveList)
		'response.End()
		
		arrApproveCode=split(strApproveList,",")
		strApproveName=""
		approve_status=""
		for i=0 to ubound(arrApproveCode)
			strApproveName=strApproveName & "," & GetUserInfo(trim(arrApproveCode(i)&""),"ENGLISHNAME")
			approve_status=approve_status & ",0"
		next
		strApproveName=mid(strApproveName,2)
		
		approve_code=strApproveList
		approve_name=strApproveName
		current_approve_code=strPRE
		current_approve_name=GetUserInfo(strPRE,"ENGLISHNAME")
		current_approve_email=GetUserInfo(strPRE,"EMAIL")
		'current_approve_email="howard.zhang@knowles.com"
		approve_status=mid(approve_status,2)
		'for OBI Test
'		approve_code=strUserCode
'		approve_name="KEM Tester"
'		current_approve_code=strUserCode
'		current_approve_name="KEM Tester"
'		current_approve_email=GetUserInfo(strUserCode,"EMAIL")
'		approve_status="0"
		
		'get act person info
		'record a new form
		SQL="select * from PROFILE_FORM where USER_CODE='"&session("code")&"' and FORM_ID='"&id&"'"
		rs.open SQL,conn,1,3
		rs.addnew
		NID="PF"&NID_SEQ("PROFILEFORM")
		rs("NID")=NID
		rs("USER_CODE")=session("code")
		rs("FORM_ID")=id
		rs("FORM_STATUS")=1 '0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=Approved;6=deny to transact
		rs("FORM_TYPE")=form_type
		rs("PARAM1")=param1
		rs("PARAM2")=param2
		rs("PARAM3")=param3
		rs("PARAM4")=param4
		rs("PARAM5")=param5
		rs("NOTE")=note
		rs("APPROVE_CODE")=approve_code
		rs("APPROVE_NAME")=approve_name
		rs("APPROVE_TIME")=approve_time
		rs("CURRENT_APPROVE_CODE")=current_approve_code
		rs("CURRENT_APPROVE_EMAIL")=current_approve_email
		rs("APPROVE_STATUS")=approve_status
		rs("APPLY_TIME")=now()
		rs("ALERT_TIME")=alert_time
		rs("ALERT_PERSON")=alert_code
		rs("ALERT_NAME")=alert_name
		rs("ALERT_EMAIL")=alert_email
		if form_type="1" then
			rs("ACTOR_CODE")=last_approve_code
			if alert_time<>"" then
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="CREATE_TASK_JOB"
				cmd.CommandType=4
				cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, "")
				cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, "")
				cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, "")
				cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, "")
				cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, "")
				cmd.Parameters.Append cmd.CreateParameter("package_name", adVarChar, adParamInput, 20, "JOB_SCRAP_ALERT")
				cmd.Parameters.Append cmd.CreateParameter("schedule_type", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("weekday", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("hour_interval", adInteger, adParamInput, 20)
				cmd.Parameters.Append cmd.CreateParameter("param1", adVarChar, adParamInput, 200, NID)
				cmd.Parameters.Append cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "")
				cmd.Parameters.Append cmd.CreateParameter("param3", adVarChar, adParamInput, 200, "")
				cmd.Parameters.Append cmd.CreateParameter("param4", adVarChar, adParamInput, 2000, "")
				cmd.Parameters.Append cmd.CreateParameter("param5", adVarChar, adParamInput, 2000, "")
				cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, "")
				cmd.Parameters.Append cmd.CreateParameter("runtime", adVarChar, adParamInput, 30, cstr(alert_time))
				cmd.Parameters.Append cmd.CreateParameter("job_id", adInteger, adParamOutput)
				cmd.execute
				oracle_job_id=cmd("job_id")
				set cmd=nothing
				rs("ORACLE_JOB_ID")=oracle_job_id
			end if
		else
		rs("ACTOR_CODE")=group_leader
		end if
		rs.update
		rs.close
		word="Successfully save new Task.成功保存报废审批新任务。"
		if fromsite="scrap" then
		action="location.href='http://"&application("ClientServerQA")&"/scrap/scrap1.asp?factory="&factory&"'"
		elseif fromsite="scrapchange" or fromsite="scrapdelete" then
		action="location.href='http://"&application("ClientServerQA")&"/scrap/jobscrap.asp?jobnumber="&param1&"'"
		elseif fromsite="board" then
		action="location.href='MyForm.asp'"
		end if
'		if form_type="1" then'send email notification to next approve
			sendEmail application("MailSender"),current_approve_email,"Please Approve "&session("user")&"'s Barcode Form. 请核准"&session("user")&"申请的系统表单。",mailhead&"<p>Hi, "&current_approve_name&"<p>Please Approve "&session("user")&"'s Barcode Form. 请核准"&session("user")&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/ApproveMyForm.asp?formid="&NID&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
			'CC to QA,PE,AME,
			set rsCC=server.CreateObject("adodb.recordset")
			SQLCC="Select U.EMAIL from USERS U inner join USERS_LINE UL on U.USER_CODE=UL.USER_CODE where U.EMAIL is not null and UL.LINE_ID=(select L.NID from LINE L inner join JOB_MASTER J on L.LINE_NAME=J.LINE_NAME where J.JOB_NUMBER='"&param1&"')"
			rsCC.open SQLCC,conn,1,3
			if not rsCC.eof then
			CC_Recievers=CC_Recievers&rsCC("EMAIL")&","
			end if
			rsCC.close
			sendEmail application("MailSender"),CC_Recievers,"Please Refer "&session("user")&"'s Barcode Form. 请参考"&session("user")&"申请的系统表单。",mailhead&"<p>Hi, "&current_approve_name&"<p>"&note&"<p>Please Refer "&session("user")&"'s Barcode Form. 请参考"&session("user")&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/MyFormSee.asp?formid="&NID&"'>"&replace(form_name,"Approve","Refer")&" "&replace(form_chinese_name,"核准","参考")&"</a><p>Best Regards!<p>"&application("SystemName")

		set rsCC=nothing
		set rst=nothing

else
		'get act person info
		'record a new form
		SQL="select * from PROFILE_FORM where USER_CODE='"&session("code")&"' and FORM_ID='"&id&"'"
		rs.open SQL,conn,1,3
		rs.addnew
		NID="PF"&NID_SEQ("PROFILEFORM")
		rs("NID")=NID
		rs("USER_CODE")=session("code")
		rs("FORM_ID")=id
		rs("FORM_STATUS")=1 '0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=Approved;6=deny to transact
		rs("FORM_TYPE")=form_type
		rs("PARAM1")=param1
		rs("PARAM2")=param2
		rs("PARAM3")=param3
		rs("PARAM4")=param4
		rs("PARAM5")=param5
		rs("NOTE")=note
'		rs("APPROVE_CODE")=approve_code
'		rs("APPROVE_NAME")=approve_name
		rs("APPROVE_CODE")="9963"
		rs("APPROVE_NAME")="KEM Tester"
		rs("APPROVE_TIME")=approve_time
		'rs("CURRENT_APPROVE_CODE")=current_approve_code
		rs("CURRENT_APPROVE_CODE")="9963"
		rs("CURRENT_APPROVE_EMAIL")=current_approve_email
		rs("APPROVE_STATUS")=approve_status
		rs("APPLY_TIME")=now()
		rs("ALERT_TIME")=alert_time
		rs("ALERT_PERSON")=alert_code
		rs("ALERT_NAME")=alert_name
		rs("ALERT_EMAIL")=alert_email
		if form_type="1" then
			rs("ACTOR_CODE")=last_approve_code
			if alert_time<>"" then
				set cmd=server.CreateObject("Adodb.Command") 
				cmd.ActiveConnection=conn 
				cmd.CommandText="CREATE_TASK_JOB"
				cmd.CommandType=4
				cmd.Parameters.Append cmd.CreateParameter("profile_task_id", adVarChar, adParamInput, 10, "")
				cmd.Parameters.Append cmd.CreateParameter("this_task_id", adVarChar, adParamInput, 10, "")
				cmd.Parameters.Append cmd.CreateParameter("task_name", adVarChar, adParamInput, 50, "")
				cmd.Parameters.Append cmd.CreateParameter("job_owner", adVarChar, adParamInput, 4, "")
				cmd.Parameters.Append cmd.CreateParameter("owner_name", adVarChar, adParamInput, 50, "")
				cmd.Parameters.Append cmd.CreateParameter("package_name", adVarChar, adParamInput, 20, "JOB_SCRAP_ALERT")
				cmd.Parameters.Append cmd.CreateParameter("schedule_type", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("weekday", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("hour_interval", adInteger, adParamInput, 20)
				cmd.Parameters.Append cmd.CreateParameter("param1", adVarChar, adParamInput, 20, NID)
				cmd.Parameters.Append cmd.CreateParameter("param2", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("param3", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("param4", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("param5", adVarChar, adParamInput, 20, "")
				cmd.Parameters.Append cmd.CreateParameter("codes", adVarChar, adParamInput, 2000, "")
				cmd.Parameters.Append cmd.CreateParameter("runtime", adVarChar, adParamInput, 30, cstr(alert_time))
				cmd.Parameters.Append cmd.CreateParameter("job_id", adInteger, adParamOutput)
				cmd.execute
				oracle_job_id=cmd("job_id")
				set cmd=nothing
				rs("ORACLE_JOB_ID")=oracle_job_id
			end if
		else
		rs("ACTOR_CODE")=group_leader
		end if
		rs.update
		rs.close
		word="Successfully save new Task.成功保存新任务。"
		if fromsite="scrap" then
		action="location.href='http://"&application("ClientServerQA")&"/scrap/scrap1.asp?factory="&factory&"'"
		elseif fromsite="scrapchange" or fromsite="scrapdelete" then
		action="location.href='http://"&application("ClientServerQA")&"/scrap/jobscrap.asp?jobnumber="&param1&"'"
		elseif fromsite="board" then
		action="location.href='MyForm.asp'"
		end if
		if form_type="1" then'send email notification to next approve
			sendEmail application("MailSender"),current_approve_email,"Please Approve "&session("user")&"'s Barcode Form. 请核准"&session("user")&"申请的系统表单。",mailhead&"<p>Hi, "&current_approve_name&"<p>Please Approve "&session("user")&"'s Barcode Form. 请核准"&session("user")&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/ApproveMyForm.asp?formid="&NID&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
		else'or send email notification to actor person if form doesn't need to be approved
			if group_leader_mail<>"" then
				sendEmail application("MailSender"),group_leader_mail,"Please transact "&session("user")&"'s Barcode Form. 请处理"&session("user")&"申请的系统表单。",mailhead&"<p>Hi, Group Leader<p>Please transact "&session("user")&"'s Barcode Form. 请处理"&session("user")&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/TransactMyForm.asp?formid="&NID&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
				if session("code")="1194" then
				SQL="update PROFILE_FORM set EMAIL_ALERT_TIMES=EMAIL_ALERT_TIMES+1 where NID='"&NID&"'"
				else
				SQL="update PROFILE_FORM set EMAIL_ALERT_TIMES=EMAIL_ALERT_TIMES+1 where USER_CODE='"&session("code")&"' and NID='"&NID&"'"
				end if
				rs.open SQL,conn,1,3
				word=word&"\nSend email notification to "&group_leader_mail&".\n已发送电子邮件通知给"&group_leader_mail&"。"
			else
				word=word&"\nNo email address for transactor, please notify him/her by telephone.\n没有处理人的邮件地址，请电话通知。"
			end if
		end if
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
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->