<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Components/ControlCommon.asp" -->
<!--#include virtual="/Components/ControlAuthority.asp" -->
<!--#include virtual="/EForm/FormHistory/MenuSettings.asp" -->
<!--#include virtual="/DefaultNavFunctions.asp" -->
<!--#include virtual="/WDCS/KESWEB20_OPEN.asp" -->
<!--#include virtual="/WDCS/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/EForm/GetFormApprover.asp" -->
<!--#include virtual="/Functions/EForm/GetCurrentApprover.asp" -->
<!--#include virtual="/Functions/EForm/GenCode.asp" -->
<!--#include virtual="/Functions/GetUser.asp" -->
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
	id=trim(rs("FORM_ID")&"")
	applicant=rs("APPLICANT_NAME")
	applicant_eng=rs("APPLICANT_ENGLISH_NAME")
	apply_date=rs("APPLY_DATE")
	approver_list=left(rs("APPROVER_CODE"),len(rs("APPROVER_CODE"))-1)
end if
rs.close

pagename="/EForm/ApproveForm/ChangeFlow.asp"
pagepara=server.URLEncode("id=" & id & "&formcode=" & formcode)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Change Flow</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css" />
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JFormatFunction.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JGetKQFunction.js" type="text/javascript"></script>
<!--#include virtual="/Language/EForm/Admin/Approveform/Lan_changeflow.asp" -->
<script language="javascript">
	function formcheck()
	{
		with(document.form1)
		{	
			if(selNewApprover.selectedIndex==0)
			{
				alert("Please select an approver!");
				selNewApprover.focus();
				return false;
			}
			var CheckFlag=false;
			for(i=0;i<2;i++)
			{
				if(rdoPosition[i].checked)
				{
						CheckFlag=true;
				}
			}
			if(!CheckFlag)
			{
				alert("Please select arrpove position!");
				return false;
			}
		}
	}
</script>
</head>

<body onLoad="language()" class="Body_Default">
<!--#include virtual="/Components/Nav.asp" -->
<table width="970" border="1" cellpadding="0" cellspacing="0" class="Border_Bark_Blue" align="center">
<form id="form1" name="form1" method="post" action="ChangeFlow1.asp" onSubmit="javascript: return formcheck();">
	<tr class="TR_Dark_Blue">
		<td colspan="2"><span id="change"></span>
		</td>
	</tr>	
	<tr>
		<td align="right" colspan="2"><span id="app"></span><font color="red">
		<%	if session("language")=0 then 
				response.Write(applicant_eng)
			else
				response.Write(applicant)
			end if%></font>&nbsp;&nbsp;<span id="appdate"></span><%=apply_date%> </td>
	</tr>	
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="code"></span> </td>
		<td><%=formcode%>&nbsp;</td>
	</tr>			
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="approver"></span></td>
		<td><%=Control_Select("selNewApprover","",getUser("OPTION",""," where USER_CODE NOT IN ('" & replace(approver_list,",","','") & "')"," ORDER BY USER_CODE","",""),"","","")%>&nbsp;</td>
	</tr>	
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="position"></span></td>
		<td>
		<%if session("language")=0 then%>
			<%=Control_Radio("rdoPosition","","0","Before Me","","")%>&nbsp;<%=Control_Radio("rdoPosition","","1","After Me","","")%>
		<%else%>
			<%=Control_Radio("rdoPosition","","0","在我之前","","")%>&nbsp;<%=Control_Radio("rdoPosition","","1","在我之后","","")%>		
		<%end if%>	
			&nbsp;</td>
	</tr>		
	
	<tr>
		<td width="15%" valign="top" class="TD_Dark_Silver"><span id="remark"></span></td><td><%= Control_Textarea("txtRemark",60,8,"")%></td>
	</tr>
    <tr>
      <td colspan="2" align="center"><%=Control_Input_Hidden("formcode","",formcode,"","")%><%=Control_SubmitButton("Submit"," Submit ",false,"")%>&nbsp;<%=Control_Button("btnReturn"," Return ","","javascript:window.location='ApproveFormAssy.asp?formcode=" & formcode & "&id=" & id & "'")%>&nbsp;<%=Control_Button("feedback","Feedback",false,"window.open('/UserProfile/FeedBack/NewFeedBack.asp?page_name="&this_path&"&response_code=1570&response_person=Eric Han','_blank')")%></td>
    </tr>
</form>
</table>

</body>
</html>
<!--#include virtual="/WDCS/KESKQ_CLOSE.asp" -->
<!--#include virtual="/WDCS/KESWEB20_CLOSE.asp" -->