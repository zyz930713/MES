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
code="3449"
arrpove_code="1295"
'FOR i=1 to 2
	NID="PF000140" & CSTR(93 + i)
	
	form_name="Approve Job Scrap" 
	form_chinese_name="核准工单报废"
	
	current_approve_email=GetUserInfo(arrpove_code,"EMAIL")
	
	ename=GetUserInfo(code,"ENGLISHNAME")
	cname=GetUserInfo(code,"NAME")
	
	current_approve_name=GetUserInfo(arrpove_code,"ENGLISHNAME")
	
	'sendEmail application("MailSender"),current_approve_email,"Please Approve " & ename &"'s Barcode Form. 请核准"& cname &"申请的系统表单。",mailhead&"<p>Hi, "&current_approve_name&"<p>Please Approve "&ename&"'s Barcode Form. 请核准"&cname&"申请的系统表单<Barcode System>。<p><a href='http://"&request.ServerVariables("HTTP_HOST")&"/Profile/MyForm/ApproveMyForm.asp?formid="&NID&"'>"&form_name&" "&form_chinese_name&"</a><p>Best Regards!<p>"&application("SystemName")
'NEXT
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->