<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Components/ControlCommon.asp" -->
<!--#include virtual="/Components/ControlAuthority.asp" -->
<!--#include virtual="/EForm/MenuSettings.asp" -->
<!--#include virtual="/DefaultNavFunctions.asp" -->
<!--#include virtual="/WDCS/KESWEB20_OPEN.asp" -->
<!--#include virtual="/WDCS/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/EForm/GetFormApprover.asp" -->
<!--#include virtual="/Functions/EForm/GetCurrentApprover.asp" -->
<!--#include virtual="/Functions/EForm/GenCode.asp" -->
<%
formcode=trim(request("formcode")&"")
CurrentApprover=GetCurrentApprover(formcode)
if formcode="" then response.Redirect("WaitingApproveList.asp")
if CurrentApprover="" or trim(CurrentApprover&"")<>trim(session("code")&"") then response.Redirect("WaitingApproveList.asp")

SQL="select * from EFORM_FORM_QUEUE where FORMCODE='" & formcode & "'"
rs.open SQL,conn,1,3
if rs.eof then
	response.Redirect("WaitingApproveList.asp")
else
	if rs("APPROVE_STATUS")<>0 then response.Redirect("WaitingApproveList.asp")
	applicant=rs("APPLICANT_NAME")
end if
rs.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Reject Form</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css" />
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JFormatFunction.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JGetKQFunction.js" type="text/javascript"></script>
<!--#include virtual="/Language/EForm/Admin/Approveform/Lan_formreject.asp" -->
</head>

<body onLoad="language()" class="Body_Default">
<!--#include virtual="/Components/Nav.asp" -->
<!--#include virtual="/iFrames/AuthorityFrame.asp" -->
<table width="998" border="1" cellpadding="0" cellspacing="0" class="Border_Bark_Blue" align="center">
<form id="form1" name="form1" method="post" action="FormReject1.asp">
	<tr class="TR_Dark_Blue">
		<td colspan="2"><span id="reject"></span></td>
	</tr>	
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="fcode"></span> </td>
		<td><%=formcode%>&nbsp;</td>
	</tr>	
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="applicant"></span> </td>
		<td><%=applicant%>&nbsp;</td>
	</tr>		
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="reason"></span></td><td><%= Control_Textarea("txtReason",60,8,"")%></td>
	</tr>
    <tr>
      <td colspan="2" align="center"><%=Control_Input_Hidden("formcode","",formcode,"","")%><%=Control_SubmitButton("Submit"," Reject ",false,"")%>&nbsp;<%=Control_Button("btnReturn"," Return ","","javascript:window.location='ApproveFormAssy.asp?formcode=" & formcode & "&id=" & id & "'")%>&nbsp;<%=Control_Button("feedback","Feedback",false,"window.open('/UserProfile/FeedBack/NewFeedBack.asp?page_name="&this_path&"&response_code=1570&response_person=Eric Han','_blank')")%></td>
    </tr>
</form>
</table>

</body>
</html>
<!--#include virtual="/WDCS/KESKQ_CLOSE.asp" -->
<!--#include virtual="/WDCS/KESWEB20_CLOSE.asp" -->