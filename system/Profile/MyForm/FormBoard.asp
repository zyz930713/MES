<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Profile/MyForm/FormFunction.asp" -->
<!--#include virtual="/Functions/GetForm.asp" -->
<!--#include virtual="/Functions/GetUserGroup.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
if session("code")="1488" then
	'session("code")="1434"
end if

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/MyForm/MyForm.asp"
formtype=request("formtype")
paramvalue=request("paramvalue")
fromdate=request("fromdate")
todate=request("todate")
if formtype<>"" then
where=" and PF.FORM_ID='"&formtype&"'"
end if
if paramvalue<>"" then
where=" and (PF.PARAM1='"&paramvalue&"' or PF.PARAM2='"&paramvalue&"' or PF.PARAM3='"&paramvalue&"' or PF.PARAM4='"&paramvalue&"')"
end if
if fromdate<>"" then
where=where&" and PF.APPLY_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
else
fromdate=dateadd("d",-7,date())
where=where&" and PF.APPLY_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and PF.APPLY_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
if session("code")="1194" then
SQL="select PF.*,U.USER_NAME,U2.USER_NAME as ACTOR_NAME,F.FORM_NAME from PROFILE_FORM PF inner join USERS U on PF.USER_CODE=U.USER_CODE left join USERS U2 on PF.ACTOR_CODE=U2.USER_CODE inner join FORM F on PF.FORM_ID=F.NID where PF.FORM_TYPE='1'  "&where&" order by APPLY_TIME desc"
else
SQL="select PF.*,U.USER_NAME,U2.USER_NAME as ACTOR_NAME,F.FORM_NAME from PROFILE_FORM PF inner join USERS U on PF.USER_CODE=U.USER_CODE left join USERS U2 on PF.ACTOR_CODE=U2.USER_CODE inner join FORM F on PF.FORM_ID=F.NID where PF.FORM_TYPE='1' and PF.CURRENT_APPROVE_CODE='"&session("code")&"' "&where&" order by APPLY_TIME desc"
end if
rs.open SQL,conn,1,3
%>
<html>
<head>
<title><%=application("SystemName")%> - Approve Form</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Profile/MyForm/Lan_MyForm.asp" -->
</head>

<body onLoad="language()">
<form name="form1" method="post" action="/Profile/MyForm/FormBoard.asp" target="_self">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#FF6600" bordercolordark="#FFFFFF">
    <tr class="t-b-midautumn">
      <td height="20" colspan="7">Search My Form </td>
    </tr>
    <tr>
      <td height="20">Form Type </td>
      <td height="20"><select name="formtype" id="formtype">
        <option value="">--Select--</option>
		<%GROUP_ID=getUserGroup("TEXT_NID",""," where U.GROUP_MEMBERS like '%"&session("code")&"%'","",",")
		if GROUP_ID<>"" then
		GROUP_ID=left(GROUP_ID,len(GROUP_ID)-1)
		end if%>
        <%= getForm("OPTION",formtype," where APPLY_GROUP in ('"&GROUP_ID&"')"," order by FORM_NAME","") %>
      </select></td>
      <td>Param Value </td>
      <td><input name="paramvalue" type="text" id="paramvalue" value="<%=paramvalue%>"></td>
      <td>Apply Time </td>
      <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script></td>
      <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
  </table>
</form>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_Browse"></span>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="13" class="t-c-greenCopy">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="white"><span id="inner_BrowseUser"></span>: <% =session("user")%>&nbsp;</td>
          <td width="50%"><div align="right"><a href="/Profile/MyForm/NewMyForm.asp" class="white"><span id="inner_Add"></span></a></div></td>
        </tr>
      </table></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20"><div align="center"><span id="inner_Action"></span></div></td>
    <td>Form ID </td>
    <td height="20"><div align="center"><span id="inner_User"></span></div></td>
    <td height="20"><div align="center"><span id="inner_FormName"></span></div></td>
    <td height="20"><div align="center"><span id="inner_FormStatus"></span></div></td>
	<td height="20"><div align="center"><span id="inner_Param1"></span></div></td>
    <td height="20"><div align="center"><span id="inner_Param2"></span></div></td>
    <td height="20"><div align="center"><span id="inner_Param3"></span></div></td>
    <td height="20"><div align="center"><span id="inner_ApproveName"></span></div></td>
    <td height="20"><div align="center"><span id="inner_ActorCode"></span></div></td>
    <td height="20"><div align="center"><span id="inner_ApplyTime"></span></div></td>
    <td height="20"><div align="center"><span id="inner_MailTimes"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof%>
  <tr class="t-c-GrayLight">
    <td height="20"><div align="center"><% =i%></div></td>
    <td height="20" bordercolorlight="#000099"><div align="center" class="red"><%if rs("FORM_STATUS")="1" then%>
    <span style="cursor:hand" onClick="javascript:window.open('ApproveMyForm.asp?formid=<%=rs("NID")%>&fromsite=board&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/Yes.gif" width="16" height="16"></span>
    <%end if%>&nbsp;</div></td>
    <td><% =rs("NID")%>&nbsp;</td>
    <td height="20"><div align="center">
      <% =rs("USER_NAME")%>
    </div></td>
	<td height="20"><div align="center"><!--<span class="red" style="cursor:hand" onClick="javascript:window.open('MyFormLog.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')">-->
	<span style="cursor:hand" onClick="javascript:window.open('ApproveMyForm.asp?formid=<%=rs("NID")%>&fromsite=board&path=<%=path%>&query=<%=query%>','main')">
      <% =rs("FORM_NAME")%></span>
    </div></td>
	<td height="20"><div align="center">
	<% =getFormStatus(rs("FORM_STATUS"))%>&nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("PARAM1")%></div></td>
    <td height="20"><div align="center">
      <% =rs("PARAM2")%>
    &nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("PARAM3")%>
    &nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("APPROVE_NAME")%>&nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("ACTOR_NAME")%>(<% =rs("ACTOR_CODE")%>)
    </div></td>
    <td><div align="center"><%=rs("APPLY_TIME")%></div></td>
    <td><div align="center"><%=rs("EMAIL_ALERT_TIMES")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="13"><div align="center"><span id="inner_Code"></span>No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->