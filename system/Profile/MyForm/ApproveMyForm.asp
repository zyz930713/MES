<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Profile/MyForm/FormFunction.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/MyForm/ApproveMyForm.asp"
formid=request.QueryString("formid")
fromsite=request.QueryString("fromsite")
SQL="select PF.*,F.FORM_NAME,F.FORM_CHINESE_NAME,F.FORM_TYPE,F.PARAM1 as F_PARAM1,F.PARAM_CHINESE1 as F_PARAM_CHINESE1,F.PARAM2 as F_PARAM2,F.PARAM_CHINESE2 as F_PARAM_CHINESE2,F.PARAM3 as F_PARAM3,F.PARAM_CHINESE3 as F_PARAM_CHINESE3,F.PARAM4 as F_PARAM4,F.PARAM_CHINESE4 as F_PARAM_CHINESE4,F.PARAM5 as F_PARAM5,F.PARAM_CHINESE5 as F_PARAM_CHINESE5,U.USER_NAME from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID inner join USERS U on PF.USER_CODE=U.USER_CODE where PF.NID='"&formid&"' and PF.CURRENT_APPROVE_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Approve System Form/核准系统表单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Profile/MyForm/ApproveFormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Profile/MyForm/Lan_ApproveMyForm.asp" -->
</head>

<body onLoad="language()">
<%if not rs.eof then
	if rs("FORM_STATUS")="1" then%>
<form name="form1" method="post" action="/Profile/MyForm/ApproveMyForm1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="19%" height="20"><span id="inner_FormName"></span></td>
    <td width="81%" height="20">
	<%if session("language")="0" then%>
	<%=rs("FORM_NAME")%>
	<%else%>
	<%=rs("FORM_CHINESE_NAME")%>
	<%end if%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="19%" height="20"><span id="inner_Applicant"></span></td>
    <td width="81%" height="20">
	<%=rs("USER_NAME")%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="19%" height="20"><span id="inner_FormType"></span></td>
    <td width="81%" height="20">
	<%=getFormType(rs("FORM_TYPE"))%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="19%" height="20"><span id="inner_FormStatus"></span></td>
    <td width="81%" height="20">
	<%=getFormStatus(rs("FORM_STATUS"))%></td>
  </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param1"></span><%=getParamName(rs("F_PARAM1"),rs("F_PARAM_CHINESE1"))%></td>
      <td height="20"><%=rs("PARAM1")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param2"></span><%=getParamName(rs("F_PARAM2"),rs("F_PARAM_CHINESE2"))%></td>
      <td height="20"><%=rs("PARAM2")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param3"></span><%=getParamName(rs("F_PARAM3"),rs("F_PARAM_CHINESE3"))%></td>
      <td height="20"><%=rs("PARAM3")%>&nbsp;</td>
    </tr>
	 <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param4"></span><%=getParamName(rs("F_PARAM4"),rs("F_PARAM_CHINESE4"))%></td>
      <td height="20"><%=rs("PARAM4")%>&nbsp;</td>
    </tr>
	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param5"></span><%=getParamName(rs("F_PARAM5"),rs("F_PARAM_CHINESE5"))%></td>
      <td height="20"><%=rs("PARAM5")%>&nbsp;</td>
    </tr>
	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_JobInfo"></span></td>
      <td height="20"><%=rs("NOTE")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ApproveFlow"></span></td>
      <td height="20"><%=replace(rs("APPROVE_NAME"),","," -> ")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ActPerson"></span></td>
      <td height="20"><%=rs("ACTOR_CODE")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ActionType"></span>&nbsp;</td>
      <td height="20"><input name="action" type="radio" value="1"><span id="inner_ActionApprove"></span>
      <input name="action" type="radio" value="2"><span id="inner_ActionDisapprove"></span></td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_DisapproveReason"></span>&nbsp;</td>
      <td height="20"><textarea name="denyreason" cols="50" rows="5" id="denyreason"></textarea></td>
    </tr>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="2"><div align="center">
      <input name="formid" type="hidden" id="formid" value="<%=formid%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
	  <input name="fromsite" type="hidden" id="fromsite" value="<%=fromsite%>">
	  <input name="Transact" type="submit" id="Transact" value="Update">
      &nbsp;
      <input name="Reset" type="reset" id="Reset" value="Reset">
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
	<%else%>
	Form has be transacted or disapproved or rejected.表单已执行过或被否决或拒绝。<br>
	<a href="/Profile/MyForm/MyFormSee.asp?formid=<%=formid%>">点击查看表单明细</a>
	<%end if
else%>
No authority to access or it is not existed.无权访问或不存在。
<%end if%>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->