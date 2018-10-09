<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/GetForm.asp" -->
<!--#include virtual="/Functions/GetUserGroup.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
form_id=request.QueryString("form_id")
jobnumber=request.QueryString("jobnumber")
pagename="/Profile/MyForm/NewMyForm.asp"
if jobnumber<>"" then
	SQL="select * from PROFILE_FORM where PARAM1='"&jobnumber&"' and FORM_STATUS=1"
	rs.open SQL,conn,1,3
	if not rs.eof then
	response.Redirect("/Profile/MyForm/EditMyForm.asp?id="&rs("NID"))
	end if
	rs.close
end if
%>
<html>
<head>
<title><%=application("SystemName")%> - Create New System Form</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Profile/MyForm/JFucntions.js" type="text/javascript"></script>
<script language="JavaScript" src="/Profile/MyForm/FormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Profile/MyForm/Lan_NewMyForm.asp" -->
</head>

<body onLoad="language();getParam('<%=jobnumber%>')">
<div id="paramDiv" style="visibility: hidden; position: absolute"><iframe id="paramFrame"></iframe></div>
<div id="jobinfoDiv" style="visibility: hidden; position: absolute"><iframe id="JobInfoFrame"></iframe></div>
<form name="form1" method="post" action="/Profile/MyForm/NewMyForm1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_FormType"></span>&nbsp;<span class="red">*</span></td>
    <td width="90%" height="20">
	<%GROUP_ID=getUserGroup("TEXT_NID",""," where U.GROUP_MEMBERS like '%"&session("code")&"%'","",",")
	if GROUP_ID<>"" then
	GROUP_ID=left(GROUP_ID,len(GROUP_ID)-1)
	end if%>
	<select name="thisform" id="thisform" onChange="getParam('')">
	<option value="">--Select--</option>
	<%= getForm("OPTION",form_id," where APPLY_GROUP in ('"&GROUP_ID&"')"," order by FORM_NAME","") %>
    </select></td>
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
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->