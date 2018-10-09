<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("PATH_INFO")
query=request.QueryString("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/Task/NewTask.asp"
id=request.QueryString("id")
if session("code")="1194" then
SQL="select PF.*,F.FORM_NAME from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID where PF.NID='"&id&"'"
else
SQL="select PF.*,F.FORM_NAME from PROFILE_FORM PF inner join FORM F on PF.FORM_ID=F.NID where PF.NID='"&id&"' and PF.USER_CODE='"&session("code")&"'"
end if
rs.open SQL,conn,1,3
paramvalue1=rs("PARAM1")
paramvalue2=rs("PARAM2")
paramvalue3=rs("PARAM3")
paramvalue4=rs("PARAM4")
paramvalue5=rs("PARAM5")
select case rs("FORM_NAME")
case "Decrease Job Quantity"
formtype_value=0 
case "Change Job Line"
formtype_value=1
end select
%>
<html>
<head>
<title>Edit My Form</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Profile/MyForm/JFucntions.js" type="text/javascript"></script>
<script language="JavaScript" src="/Profile/MyForm/FormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Profile/MyForm/Lan_EditMyForm.asp" -->
</head>

<body onLoad="language();getEditParam();">
<%if not rs.eof then%>
<div id="paramDiv" style="visibility:hidden; position:absolute"><iframe id="paramFrame"></iframe></div>
<div id="jobinfoDiv" style="visibility:hidden; position:absolute"><iframe id="JobInfoFrame"></iframe></div>
<form name="form1" method="post" action="/Profile/MyForm/EditMyForm1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_FormType"></span>&nbsp;<span class="red">*</span></td>
    <td width="90%" height="20"><%=rs("FORM_NAME")%></td>
  </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param1"></span></td>
      <td height="20"><span id="paramhtml1"></span>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param2"></span></td>
      <td height="20"><span id="paramhtml2"></span>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param3"></span></td>
      <td height="20"><span id="paramhtml3"></span>&nbsp;</td>
    </tr>
	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param4"></span></td>
      <td height="20"><span id="paramhtml4"></span>&nbsp;</td>
    </tr>
		<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Param5"></span></td>
      <td height="20"><span id="paramhtml5"></span>&nbsp;</td>
    </tr>
  	<tr class="t-c-GrayLight">
      <td height="20"><span id="inner_JobInfo"></span></td>
      <td height="20"><span id="jobinfohtml"></span>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ApproveFlow"></span></td>
      <td height="20"><span id="approveflow"></span>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_ActPerson"></span></td>
      <td height="20"><span id="actperson"></span>&nbsp;</td>
    </tr>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="2"><div align="center">
      <input name="formid" type="hidden" id="formid" value="<%=rs("FORM_ID")%>">
	  <input name="paramvalue1" type="hidden" id="paramvalue1" value="<%=rs("PARAM1")%>">
	  <input name="paramvalue2" type="hidden" id="paramvalue2" value="<%=rs("PARAM2")%>">
	  <input name="paramvalue3" type="hidden" id="paramvalue3" value="<%=rs("PARAM3")%>">
	  <input name="paramvalue4" type="hidden" id="paramvalue4" value="<%=rs("PARAM4")%>">
  	  <input name="paramvalue5" type="hidden" id="paramvalue5" value="<%=rs("PARAM5")%>">
      <input name="profile_form_id" type="hidden" id="profile_form_id" value="<%=id%>">
	  <input name="line_name" type="hidden" id="line_name" value="">
	  <input name="start_quantity" type="hidden" id="start_quantity" value="">
	  <input name="group_leader" type="hidden" id="group_leader" value="">
	  <input name="group_leader_mail" type="hidden" id="group_leader_mail" value="">
	  <input name="note" type="hidden" id="note" value="">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
	  <input name="Update" type="submit" id="Update" value="Update">
      &nbsp;
      <input name="Reset" type="reset" id="Reset" value="Reset">
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
<%end if%>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->