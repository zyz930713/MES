<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/MailHead.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
profile_form_id=request.Form("profile_form_id")
param1=trim(request.Form("param1"))
param2=trim(request.Form("param2"))
param3=trim(request.Form("param3"))
param4=trim(request.Form("param4"))
group_leader=trim(request.Form("group_leader"))
group_leader_mail=trim(request.Form("group_leader_mail"))

'record a new form
if session("code")="1194" then
SQL="select * from PROFILE_FORM where NID='"&profile_form_id&"'"
else
SQL="select * from PROFILE_FORM where USER_CODE='"&session("code")&"' and NID='"&profile_form_id&"'"
end if
rs.open SQL,conn,1,3
if not rs.eof then
form_id=rs("FORM_ID")
if session("code")<>"1194" then
rs("USER_CODE")=session("code")
end if
rs("FORM_STATUS")=1
rs("PARAM1")=request.Form("param1")
rs("PARAM2")=request.Form("param2")
rs("PARAM3")=request.Form("param3")
rs("PARAM4")=request.Form("param4")
rs("NOTE")=request.Form("note")
rs("APPROVE_CODE")=""
rs("APPROVE_NAME")=""
rs("APPROVE_TIME")=""
rs("CURRENT_APPROVE_CODE")=""
rs("APPROVE_STATUS")=""
rs("APPLY_TIME")=now()
rs("ACTOR_CODE")=group_leader
rs.update
end if
rs.close

'get info about form
SQL="select * from FORM where NID='"&form_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
packagename=rs("PACKAGE")
form_name=rs("FORM_NAME")
form_chinese_name=rs("FORM_CHINESE_NAME")
form_type=rs("FORM_TYPE")
end if
rs.close

word="Successfully edit new Task.成功修改新任务。"
action="location.href='MyForm.asp'"
if form_type="1" then'send email notification to next approve
	
else'or send email notification to actor person if form doesn't need to be approved
	if group_leader_mail<>"" then
		set JMail=CreateObject("JMail.Message") 
		  JMail.ContentType = "text/html"
		  JMail.Charset ="gb2312"
		  JMail.From = "BarcodeSystem@knowles.com" 
		  JMail.AddRecipient group_leader_mail
		  'JMail.AddRecipient "dickens.xu@knowles.com"
		  JMail.Subject = "Please transact "&session("user")&"'s Barcode Form. 请处理"&session("user")&"申请的系统表单。"
		  JMail.Body = mailhead&"<p>Hi, Group Leader<p>Please transact "&session("user")&"'s Barcode Form.请处理"&session("user")&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/TransactMyForm.asp?formid="&profile_form_id&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
		  JMail.Send (application("MailServer"))
		set JMail = Nothing
		if session("code")="1194" then
		SQL="update PROFILE_FORM set EMAIL_ALERT_TIMES=EMAIL_ALERT_TIMES+1 where NID='"&profile_form_id&"'"
		else
		SQL="update PROFILE_FORM set EMAIL_ALERT_TIMES=EMAIL_ALERT_TIMES+1 where USER_CODE='"&session("code")&"' and NID='"&profile_form_id&"'"
		end if
		rs.open SQL,conn,1,3
		word=word&"\nSend email notification to "&group_leader_mail&".\n已发送电子邮件通知给"&group_leader_mail&"。"
	else
		word=word&"\nNo email address for transactor, please notify him/her by telephone.\n没有处理人的邮件地址，请电话通知。"
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