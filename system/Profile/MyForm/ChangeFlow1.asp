<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WDCS/KESWEB20_OPEN.asp" -->
<!--#include virtual="/WDCS/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/GetUserEmail.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/EForm/GetFormContent.asp" -->
<!--#include virtual="/Functions/GetPageHTML.asp" -->
<!--#include virtual="/Functions/EForm/GetEFormName.asp" -->
<%
formcode=trim(request("formcode")&"")
newApprover=trim(request("selNewApprover")&"")
strPosition=trim(request("rdoPosition")&"")
strRemark=trim(request("txtRemark")&"")
if formcode="" then response.Redirect("WaitingApproveList.asp")
SQL="select * from EFORM_FORM_QUEUE where FORMCODE='" & formcode & "'"
rs.open SQL,conn,1,3

if rs.eof then
	response.Redirect("WaitingApproveList.asp")
else
	formid=trim(rs("FORM_ID")&"")
	ApproverList=trim(rs("APPROVER_CODE")&"")
	if strPosition=1 then
		ApproverList=replace(ApproverList,session("code"),session("code")&","&newApprover)
		
	else
		ApproverList=replace(ApproverList,session("code"),newApprover&","&session("code"))
		
		rs("CURRENT_APPROVER")=newApprover
	end if
	arrCodes=split(ApproverList,",")
	y=ubound(arrCodes)
	if trim(arrCodes(y)&"")<>"" then
		strLastApprover=trim(arrCodes(y)&"")
	else
		strLastApprover=trim(arrCodes(y-1)&"")
	end if
	if right(ApproverList,1)<>"," then ApproverList=ApproverList&","
	intGrade=rs("APPROVE_GRADE")
	if trim(intGrade&"")="" then intGrade=0
	rs("APPROVE_GRADE")=cint(intGrade)+1
	rs("APPROVER_CODE")=ApproverList
	rs("LAST_APPROVER")=strLastApprover
	rs("CHANGE_FLOW_REMARK")=strRemark
	rs("CHANGE_FLOW_CODE")=session("code")
	rs("CHANGE_FLOW_DATE")=NOW	
	rs.update
	if strPosition=0 then	
		strHTML=GetPageHTML("http://kesweba:8900/EForm/ApproveForm/ApproveFormAssy.asp?isMail=Y&id=" & formid & "&formcode=" & formcode)
		strHTML=replace(strHTML,"/CSS/","http://kesweba:8900/CSS/")
strSubject = "Please approve "& GetEFormNameByID(trim(rs("FORM_ID")&""),"ENGLISH_NAME") & " "&formcode&" of "&rs("APPLICANT_ENGLISH_NAME")&" 请您核准" &  GetEFormNameByID(trim(rs("FORM_ID")&""),"CHINESE_NAME") & " "& formcode & "申请单！"	
		strSubject = "Please approve "& GetEFormNameByID(trim(rs("FORM_ID")&""),"ENGLISH_NAME") & " "&rs("formcode")&" of "&rs("APPLICANT_ENGLISH_NAME")&" 请您核准" &  GetEFormNameByID(trim(rs("FORM_ID")&""),"CHINESE_NAME") & " "& rs("formcode") & "申请单！"
		strFrom="FormSystem@knowles.com" 

		strTo=GetUserEmail(trim(newApprover&""))
		if trim(strTo&"")<>"" then
			if instr(strTo,"@")>1 then
				SendJMail strFrom,strTo,strSubject,strHTML
			end if
		end if
	end if
	
end if
rs.close
%>
<!--#include virtual="/WDCS/KESKQ_CLOSE.asp" -->
<!--#include virtual="/WDCS/KESWEB20_CLOSE.asp" -->
<script language="javascript">
<%if session("language")=0 then%>
	alert("Approver flow for form <%=formcode%> has been changed！");
<%else%>
	alert("表单<%=formcode%>签核流程已被修改！");
<%end if%>
	window.location="WaitingApproveList.asp";
</script>

