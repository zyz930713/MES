<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
user=request.QueryString("user")
task_name=request.QueryString("task_name")
if user<>"" then
where=where&" and lower(U.USER_NAME) like '%"&lcase(user)&"%'"
end if
if task_name<>"" then
where=where&" and lower(T.TASK_NAME) like '%"&lcase(task_name)&"%'"
nwhere=nwhere&" and lower(T.TASK_NAME) like '%"&lcase(task_name)&"%'"
end if
pagename="/Profile/MyTask/MyTask.asp"
if lcase(user)="all" then
SQL="select P.*,U.USER_NAME,T.TASK_NAME as TASK_TAG,UJ.NEXT_DATE from PROFILE_TASK P inner join USERS U on P.USER_CODE=U.USER_CODE inner join TASK T on P.TASK_ID=T.NID left join USER_JOBS UJ on P.ORACLE_JOB_ID=UJ.JOB where P.NID is not null "&nwhere
else
SQL="select P.*,U.USER_NAME,T.TASK_NAME as TASK_TAG,UJ.NEXT_DATE from PROFILE_TASK P inner join USERS U on P.USER_CODE=U.USER_CODE inner join TASK T on P.TASK_ID=T.NID left join USER_JOBS UJ on P.ORACLE_JOB_ID=UJ.JOB where P.USER_CODE='"&session("code")&"'"&nwhere
end if
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Profile/MyTask/Lan_MyTask.asp" -->
</head>

<body onLoad="language()">
<form action="/Profile/MyTask/MyTask.asp" method="get" name="form1" target="_self">
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-yellowCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td class="today"><span id="inner_SearchUser"></span></td>
    <td><input name="user" type="text" id="user" value="<%=user%>"></td>
    <td class="today"><span id="inner_SearchName"></span></td>
    <td><label>
    <input name="task_name" type="text" id="task_name" value="<%=task_name%>">
    </label></td>
    <td><input name="searchbutton" type="submit" class="t-b-Yellow" id="searchbutton" value="Submit">    </td>
  </tr>
</table>
</form>
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_Browse"></span>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="12" class="t-c-greenCopy">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="white"><span id="inner_BrowseUser"></span>: <% =session("user")%>&nbsp;</td>
          <td width="50%"><div align="right"><a href="/Profile/MyTask/NewMyTask.asp" class="white"><span id="inner_Add"></span></a></div></td>
        </tr>
      </table></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20" colspan="2"><div align="center"><span id="inner_Action"></span></div></td>
    <td height="20"><div align="center"><span id="inner_User"></span></div></td>
    <td height="20"><div align="center"><span id="inner_TaskName"></span></div></td>
    <td height="20"><div align="center"><span id="inner_TaskType"></span></div></td>
    <td height="20"><div align="center"><span id="inner_ScheduleType"></span></div></td>
    <td height="20"><div align="center"><span id="inner_StartDay"></span></div></td>
    <td height="20"><div align="center"><span id="inner_RunTime"></span></div></td>
    <td height="20"><div align="center"><span id="inner_Schedule"></span></div></td>
    <td height="20"><div align="center"><span id="inner_NextRunTime"></span></div></td>
    <td height="20"><div align="center"><span id="inner_Recievers"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof%>
  <tr class="t-c-GrayLight">
    <td height="20"><div align="center"><% =i%></div></td>
    <td height="20" bordercolorlight="#000099"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditMyTask.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20" bordercolorlight="#000099" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if(confirm('Are you sure to delete this task?')){window.open('DeleteMyTask.asp?id=<%=rs("NID")%>&taskname=<%=rs("TASK_NAME")%>&job_id=<%=rs("ORACLE_JOB_ID")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
    <td height="20"><div align="center">
      <% =rs("USER_NAME")%>
    </div></td>
    <td height="20"><div align="center">
      <span class="red" style="cursor:hand" onClick="javascript:window.open('MyTaskLog.asp?id=<%=rs("NID")%>&task_id=<%=rs("TASK_ID")%>&path=<%=path%>&query=<%=query%>','main')"><% =replace(rs("TASK_NAME"),"^","'")%></span>
    </div></td>
    <td height="20"><div align="center">
      <% =rs("TASK_TAG")%>
    </div></td>
    <td height="20"><div align="center">
      <% =rs("SCHEDULE_TYPE")%>
    &nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("START_TIME")%>
    &nbsp;</div></td>
    <td height="20"><div align="center">
      <% =rs("HAPPEN_TIME")%>
    &nbsp;</div></td>
    <td height="20"><div align="center">
	<% select case rs("WEEK_DAY")
	case "2"
	week_day="Monday"
	case "3"
	week_day="Tuesday"
	case "4"
	week_day="Wendesday"
	case "5"
	week_day="Thursday"
	case "6"
	week_day="Friday"
	case "7"
	week_day="Saturday"
	case "1"
	week_day="Sunday"
	end select%>
	<%=week_day%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=rs("NEXT_DATE")%>&nbsp;</div></td>
    <td height="20"><div align="center">
      <% =formatlongstringbreak(rs("RECIEVERS"),"<br>",20)%>
    </div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="12"><div align="center"><span id="inner_Code"></span>No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->