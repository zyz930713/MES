<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WDCS/KESWEB20_OPEN.asp" -->
<!--#include virtual="/WDCS/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/GetUserEmail.asp" -->
<!--#include virtual="/Functions/SendEmailBySQL.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/EForm/GetEFormName.asp" -->
<!--#include virtual="/Functions/GetPageHTML.asp" -->
<!--#include virtual="/Functions/EForm/GetCurrentApprover.asp" -->
<%
formcode=trim(request("formcode")&"")
if formcode="" then response.Redirect("WaitingApproveList.asp")
CurrentApprover=GetCurrentApprover(formcode)
if CurrentApprover="" or trim(CurrentApprover&"")<>trim(session("code")&"") then response.Redirect("WaitingApproveList.asp")

SQL="select * from EFORM_FORM_QUEUE where FORMCODE='" & formcode & "'"
rs.open SQL,conn,1,3
if rs.eof then
	response.Redirect("WaitingApproveList.asp")
else
	if rs("APPROVE_STATUS")<>0 then response.Redirect("WaitingApproveList.asp")
	approve_step=trim(rs("APPROVE_STEP")&"")
	if approve_step="" then approve_step=0
	rs("DENY_REASON")=request("txtReason")
	rs("PASS_CODE")=rs("PASS_CODE") & session("code") & ","
	rs("PASS_NAME")=rs("PASS_NAME") & session("user") & ","	
	rs("APPROVE_DATE")=rs("APPROVE_DATE") & date & ","
	rs("APPROVE_STEP")=approve_step+1
	rs("APPROVE_STATUS")="2"
	rs.update
		form_chinese_name=GetEFormNameByID(trim(rs("FORM_ID")&""),"CHINESE_NAME")
		form_english_name=GetEFormNameByID(trim(rs("FORM_ID")&""),"ENGLISH_NAME")
		HTML="<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN'><html><head><title>"
		HTML=HTML&form_chinese_name&"表单否决通知<br>Reject notification of "
		HTML=HTML&form_english_name&"</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><link href='http://kesweba:8900/Css/general.css' rel='stylesheet' type='text/css'></head><body><table width='500' height='242' border='1' align='left' cellpadding='0' cellspacing='0' bordercolorlight='#9966CC' bordercolordark='#FFFFFF' bgcolor='#FFFFFF'><tr><td height='20' colspan='4' valign='middle' class='TD_Dark_Blue'>&nbsp;</td></tr><tr ><td height='20' valign='middle'><div align='center'><font color='#000000' size='2'>"
		HTML=HTML&form_chinese_name&"表单否决通知<br>Reject notification of "
		HTML=HTML&rs("formcode")&"</font></div></td></tr><tr ><td height='20' colspan='4'><div align='left'><font color='#000000' size='2'>您在 "
		HTML=HTML&rs("apply_date")&" 申请编号为 "
		HTML=HTML&"<a href='http://kesweba:8900/EForm/FormHistory/FormDetail.asp?id="&rs("FORM_ID")&"&formcode="&formcode&"'>"&formcode&"</a> 的"
		HTML=HTML&form_chinese_name&"没有被批准!<br>The form <a href='http://kes-web/EForm/FormHistory/FormDetail.asp?id="&rs("FORM_ID")&"&formcode="&rs("formcode")&"'>"&formcode&"</a> you applied on "
		HTML=HTML&rs("apply_date")&" has been rejected.</font></div></td></tr><tr><td height='20'><div align='left'><font color='#000000' size='2'>请到<a href='http://kes-web/EForm/FormHistory/'>追寻表单</a>中查阅该表单核准情况！<br>Please check the status of this form in <a href='http://kes-web/EForm/FormHistory/'>Pursuit form</a>.</font></div></td></tr><tr><td height='20'><div align='left'><font color='#FF0000' size='2'>本邮件内容仅为楼氏内部网站电子表单系统的邮件版，一切内容的真实性以网站电子表单系统为准。<br>All information in this email is the duplication of E-form in KES intranet web and subject to authenticity in KES internat web.</font></div></td></tr></table></body></html>"
		
		strSubject = "Your "& GetEFormNameByID(trim(rs("FORM_ID")&""),"ENGLISH_NAME") & " "&rs("formcode")&" is rejected. 您的" &  GetEFormNameByID(trim(rs("FORM_ID")&""),"CHINESE_NAME") & " "& rs("formcode") & "没有被批准！" 
		strFrom="FormSystem@knowles.com" 
		strTo=GetUserEmail(trim(rs("INPUT_CODE")&""))
		'strTo="eric.han@knowles.com"
		if trim(strTo&"")<>"" then
			if instr(strTo,"@")>1 then		
				SendJMail strFrom,strTo,strSubject,HTML
			end if
		end if
	
end if
rs.close
%>
<!--#include virtual="/WDCS/KESKQ_CLOSE.asp" -->
<!--#include virtual="/WDCS/KESWEB20_CLOSE.asp" -->
<script language="javascript">
	<%if session("language")=0 then%>
		alert("Form <%=formcode%> has been rejected！");
	<%else%>
		alert("表单<%=formcode%>已被否决！");
	<%end if%>
	window.location="WaitingApproveList.asp";
</script>

