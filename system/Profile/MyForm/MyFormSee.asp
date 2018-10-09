<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Profile/MyForm/FormFunction.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/MyForm/NewMyForm.asp"
formid=request.QueryString("formid")
SQL="select 1,PF.*,F.FORM_NAME,F.FORM_CHINESE_NAME,F.FORM_TYPE,U.USER_NAME from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID inner join USERS U on PF.USER_CODE=U.USER_CODE where PF.NID='"&formid&"'"
' and PF.ACTOR_CODE like '%"&session("Code")&"%'
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Transact System Form/处理系统表单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Profile/MyForm/Lan_MyFormSee.asp" -->
</head>

<body onLoad="language()">
<%if not rs.eof then%>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_FormName"></span></td>
    <td width="90%" height="20">
	<%if session("language")="0" then%>
	<%=rs("FORM_NAME")%>
	<%else%>
	<%=rs("FORM_CHINESE_NAME")%>
	<%end if%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_Applicant"></span></td>
    <td width="90%" height="20">
	<%=rs("USER_NAME")%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_FormType"></span></td>
    <td width="90%" height="20">
	<%=getFormType(rs("FORM_TYPE"))%></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_FormStatus"></span></td>
    <td width="90%" height="20">
	<%=getFormStatus(rs("FORM_STATUS"))%></td>
  </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param1"></span></td>
      <td height="20"><%=rs("PARAM1")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param2"></span></td>
      <td height="20"><%=rs("PARAM2")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param3"></span></td>
      <td height="20"><%=rs("PARAM3")%>&nbsp;</td>
    </tr>
	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param4"></span></td>
      <td height="20"><%=rs("PARAM4")%>&nbsp;</td>
    </tr>
	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_JobInfo"></span></td>
      <td height="20"><%=rs("NOTE")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ApproveFlow"></span></td>
      <td height="20"><%=rs("APPROVE_NAME")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ActPerson"></span></td>
      <td height="20"><%=rs("ACTOR_CODE")%>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_DenyReason"></span>&nbsp;</td>
      <td height="20"><%=rs("DENy_REASON")%>&nbsp;</td>
    </tr>
	 <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_RejectReason"></span>&nbsp;</td>
      <td height="20"><%=rs("REJECT_REASON")%>&nbsp;</td>
    </tr>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="2"><div align="center">&nbsp;
      <input name="Close" type="button" id="Reset" value="Close" onClick="javascript:window.close()">
    </div></td>
    </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
<%else%>
No authority to access or it is not existed.无权访问或不存在。
<%end if%>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->