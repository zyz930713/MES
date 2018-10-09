<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/MailHead.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/GetUserInfo.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
formid=request.Form("formid")
denyreason=request.Form("denyreason")
fromsite=request.Form("fromsite")
this_approve_status=request.Form("action")
SQL="select PF.*,F.FORM_NAME,F.FORM_CHINESE_NAME,U.USER_NAME,U.EMAIL from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID inner join USERS U on PF.USER_CODE=U.USER_CODE where PF.NID='"&formid&"' and PF.CURRENT_APPROVE_CODE='"&session("Code")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
applicant_name=rs("USER_NAME")
applicant_mail=trim(rs("EMAIL"))
form_name=rs("FORM_NAME")
form_chinese_name=rs("FORM_CHINESE_NAME")
end if
rs.close
SQL="select PF.* from PROFILE_FORM PF where PF.NID='"&formid&"' and PF.APPROVE_CODE like '%"&session("Code")&"%'"
rs.open SQL,conn,1,3
if not rs.eof then
	id=rs("FORM_ID")
	param1=rs("PARAM1")
	param2=rs("PARAM2")
	param3=rs("PARAM3")
	param4=rs("PARAM4")
	param5=rs("PARAM5")
	a_approve_code=split(trim(rs("APPROVE_CODE")),",")
	a_approve_status=split(trim(rs("APPROVE_STATUS")),",")
	'a_approve_time=split(rs("APPROVE_TIME"),",")
	approve_status=""
	approve_time=rs("APPROVE_TIME") & "," & approve_time

	for i=0 to ubound(a_approve_code)
		if session("code")=a_approve_code(i) then
			if this_approve_status="1" then 'approve
			a_approve_status(i)="2"
			else
			a_approve_status(i)="3" 'disapprove
			rs("DENY_REASON")=denyreason
			rs("FORM_STATUS")=3
			end if
			'a_approve_time(i)=now()
		end if
		approve_status=approve_status&a_approve_status(i)&","
		'approve_time=approve_time&a_approve_time(i)&","
	next
	approve_status=left(approve_status,len(approve_status)-1)
	'approve_time=mid(approve_time,2)
	rs("APPROVE_STATUS")=approve_status
	
	if trim(rs("APPROVE_TIME")&"")="" then
		rs("APPROVE_TIME")=now
	else
		rs("APPROVE_TIME")=rs("APPROVE_TIME") & "," & now
	end if
	if session("code")=a_approve_code(ubound(a_approve_code)) and this_approve_status="1" then
	rs("FORM_STATUS")=5 'form is closed
	form_status=5
	elseif session("code")<>a_approve_code(ubound(a_approve_code)) and this_approve_status="1" then
		for i=0 to ubound(a_approve_code)
			if session("code")=a_approve_code(i) and i<ubound(a_approve_code) then
			next_approve_code=a_approve_code(i+1)
			'next_approve_email=a_approve_email(i+1)
			next_approve_email=GetUserInfo(trim(next_approve_code&""),"EMAIL")
			end if
		next
	rs("CURRENT_APPROVE_CODE")=next_approve_code 'have approved
		if id="FO00000006" then
			sendEmail application("MailSender"),next_approve_email,"Please transact "&applicant_name&"'s Barcode Form. 请处理"&applicant_name&"申请的系统表单。",mailhead&"<p>Hi, " & GetUserInfo(next_approve_code,"ENGLISHNAME") & "<p>Please transact "&applicant_name&"'s Barcode Form. 请处理"&applicant_name&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/ApproveMyForm.asp?formid="&formid&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
			word="Success to approve form. System sent email notification to next approvor "&applicant_mail&"."
			cword="本次核准步骤成功。已发送电子邮件通知下一个核准人"&applicant_mail&"。"		
		else
			sendEmail application("MailSender"),next_approve_email,"Please transact "&applicant_name&"'s Barcode Form. 请处理"&applicant_name&"申请的系统表单。",mailhead&"<p>Hi, " & GetUserInfo(next_approve_code,"ENGLISHNAME") & "<p>Please transact "&applicant_name&"'s Barcode Form. 请处理"&applicant_name&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/TransactMyForm.asp?formid="&formid&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
			word="Success to approve form. System sent email notification to next approvor "&applicant_mail&"."
			cword="本次核准步骤成功。已发送电子邮件通知下一个核准人"&applicant_mail&"。"
		end if
	else
	sendEmail application("MailSender"),applicant_mail,"Your form <"&form_name&"> is rejected to be transacted by "&session("user")&".你申请的表单<"&form_chinese_name&">被"&session("user")&"否决。",mailhead&"<p>Hi, "&applicant_name&"<p>Your form <"&form_name&"> is rejected to be transacted by "&session("user")&".你申请的表单<"&form_chinese_name&">被"&session("user")&"否决。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/MyFormSee.asp?formid="&formid&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
	word="Success to deny form. System sent email notification to "&applicant_mail&"."
	cword="表单被否决。已发送电子邮件通知申请人"&applicant_mail&"。"
	end if
	rs.update
else
	word="No authority to access or it is not existed."
	cword="无权访问或不存在。"
end if

rs.close
if form_status=5 then
'auto transact form
response.Redirect("TransactMyForm1.asp?formid="&formid&"&action=1&fromsite="&fromsite)
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Form Transaction/表单处理</title>
</head>

<body>
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<td height="20"><%if session("language")="0" then%><%=word%><%else%><%=word&cword%><%end if%></td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <tr>
	<td height="20"><div align="center">
	  <input type="button" name="Button" value="<%if session("language")="0" then%>Close<%else%>关闭<%end if%>" onClick="javascript:window.close()">
	</div></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->