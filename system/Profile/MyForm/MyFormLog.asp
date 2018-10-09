<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
id=request.QueryString("id")
user=request.QueryString("user")
result=request.QueryString("result")
fromdate=request.QueryString("fromdate")
todate=request.QueryString("todate")
if user<>"" then
where=where&" and lower(U.USER_NAME) like '%"&lcase(user)&"%'"
end if
if result<>"" then
	if result="1" then
	where=where&" and L.TRANSACT_RESULT=1"
	else
	where=where&" and L.TRANSACT_RESULT=0"
	end if
end if
if fromdate<>"" then
where=where&" and L.TRANSACT_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and L.TRANSACT_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagename="/Profile/Basic/Basic.asp"
SQL="select L.*,U.USER_NAME,F.FORM_NAME from PROFILE_FORM_LOG L inner join PROFILE_FORM P on L.PROFILE_FORM_NID=P.NID inner join FORM F on P.FORM_ID=F.NID inner join USERS U on P.USER_CODE=U.USER_CODE where P.NID='"&id&"' "&where&" order by TRANSACT_TIME desc"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Profile/MyForm/Lan_MyFormLog.asp" -->
</head>

<body onLoad="language()">
<form name="form1" method="get" action="/Profile/MyForm/MyFormLog.asp">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="7" class="t-c-yellowCopy"><span id="inner_Search"></span></td>
    </tr>
    <tr>
      <td class="today"><span id="inner_SearchUser"></span></td>
      <td><input name="user" type="text" id="user" value="<%=user%>"></td>
      <td class="today"><span id="inner_SearchResult"></span></td>
      <td><label>
        <select name="result" id="result">
          <option value="">--Select--</option>
          <option value="1" <%if result="1" then%>selected<%end if%>>Success</option>
          <option value="0" <%if result="0" then%>selected<%end if%>>Fail</option>
        </select>
      </label></td>
      <td class="today"><span id="inner_SearchRuntime"></span></td>
      <td><span id="inner_SearchFrom"></span>
          <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
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
                        </script>
        &nbsp; </td>
      <td><input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="task_id" type="hidden" id="task_id" value="<%=task_id%>">
          <input name="searchbutton" type="submit" class="t-b-Yellow" id="searchbutton" value="Submit">
      </td>
    </tr>
  </table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span class="white"><span id="inner_BrowseUser"></span>:
    <% =session("user")%>&nbsp;</span></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20"><div align="center"><span id="inner_User"></span></div></td>
    <td height="20"><div align="center"><span id="inner_FormName"></span></div></td>
    <td height="20"><div align="center"><span id="inner_TransactCode"></span></div></td>
    <td height="20"><div align="center"><span id="inner_TransactName"></span></div></td>
	<td height="20"><div align="center"><span id="inner_TransactResult"></span></div></td>
    <td height="20"><div align="center"><span id="inner_TransactTime"></span></div></td>
    <td height="20"><div align="center"><span id="inner_RunErrors"></span></div></td>
    <td height="20"><div align="center"><span id="inner_ResultNote"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof%>
  <tr class="t-c-GrayLight">
    <td height="20"><div align="center"><% =i%></div></td>
    <td height="20"><div align="center">
      <% =rs("USER_NAME")%>
    </div></td>
    <td height="20"><div align="center">
      <% =rs("FORM_NAME")%>
    </div></td>
    <td height="20"><div align="center">
      <% =rs("TRANSACT_CODE")%>
    &nbsp;</div></td>
	<td height="20"><div align="center">
      <% =rs("TRANSACT_NAME")%>
    </div></td>
    <td height="20"><div align="center"><%if rs("TRANSACT_RESULT")="1" then%>OK<%else%>Fail<%end if%></div></td>
    <td height="20"><div align="center">
      <% =rs("TRANSACT_TIME")%>
    &nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("ERROR_INFO")%>
    &nbsp;</div></td>
    <td><% =rs("RESULT_NOTE")%></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="9"><div align="center"><span id="inner_Code"></span>No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->