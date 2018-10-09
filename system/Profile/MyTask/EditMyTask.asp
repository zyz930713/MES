<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<!--#include virtual="/Functions/GetTask.asp" -->
<!--#include virtual="/System/AdminCheck.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/Task/NewTask.asp"
id=request.QueryString("id")
if sysadmin=true then
SQL="select PT.*,T.TASK_NAME as TASK_TAG from PROFILE_TASK PT inner join TASK T on PT.TASK_ID=T.NID where PT.NID='"&id&"'"
else
SQL="select PT.*,T.TASK_NAME as TASK_TAG from PROFILE_TASK PT inner join TASK T on PT.TASK_ID=T.NID where PT.NID='"&id&"' and PT.USER_CODE='"&session("code")&"'"
end if
rs.open SQL,conn,1,3
if not rs.eof then
paramvalue1=rs("PARAM1")
paramvalue2=rs("PARAM2")
paramvalue3=rs("PARAM3")
paramvalue4=rs("PARAM4")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Profile/MyTask/FormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Profile/MyTask/Lan_EditMyTask.asp" -->
<script language="javascript">
function getParam()
{
var task_id="<%=rs("TASK_ID")%>";
	with(document.form1)
	{
		if(task_id!="")
		{
			document.all.paramFrame.src="/Functions/getTaskParam.asp?id="+task_id+"&paramvalue1=<%=paramvalue1%>&paramvalue2=<%=paramvalue2%>&paramvalue3=<%=paramvalue3%>&paramvalue4=<%=paramvalue4%>";
		}
		else
		{
			document.all.paramhtml1.innerHTML="";
			document.all.paramhtml2.innerHTML="";
			document.all.paramhtml3.innerHTML="";
			document.all.paramhtml4.innerHTML="";
		}
	}
}
function isperiod()
{
	with(document.form1)
	{
		if(period.checked)
		{
			happenitem.disabled=false;
			fromdate.disabled=false;
			happentime1.disabled=false;
			happentime2.disabled=false;
		}
		else
		{
			happenitem.disabled=true;
			fromdate.disabled=true;
			happentime1.disabled=true;
			happentime2.disabled=true;
			happenitem.options[0].selected=true;
		}
		isweek()
	}
}
function isweek()
{
	with(document.form1)
	{
		if(happenitem.selectedIndex==3)
		{
			span_hour.style.visibility="hidden";
			span_week.style.visibility="visible";
			week[0].disabled=false;
			week[1].disabled=false;
			week[2].disabled=false;
			week[3].disabled=false;
			week[4].disabled=false;
			week[5].disabled=false;
			week[6].disabled=false;
		}
		else if(happenitem.selectedIndex==1)
		{
			span_hour.style.visibility="visible";
			span_week.style.visibility="hidden";
			week[0].disabled=true;
			week[1].disabled=true;
			week[2].disabled=true;
			week[3].disabled=true;
			week[4].disabled=true;
			week[5].disabled=true;
			week[6].disabled=true;
		}
		else
		{
			span_hour.style.visibility="hidden";
			span_week.style.visibility="hidden";
		}
	}
}
function time1()
{
	with(document.form1)
	{
		if (!(isNumberString(happentime1.value,"1234567890"))||new Number(happentime1.value)>24)
		{
			alert("时间格式错误！");
			happentime1.value="00";
		}	
		else
		{
			if (happentime1.value=="24")
			{
			happentime1.value="00";
			}
		}
	}
}
function time2()
{
	with(document.form1)
	{
		if (!(isNumberString(happentime2.value,"1234567890"))||new Number(happentime2.value)>60)
		{
			alert("时间格式错误！");
			happentime2.value="00";
		}	
		else
		{
			if (happentime2.value=="60")
			{
			happentime2.value="00";
			}
		}
	}
}
</script>
</head>

<body onLoad="language();isperiod();isweek();getParam()">
<div id="paramDiv" style="visibility:hidden; position:absolute"><iframe id="paramFrame"></iframe></div>
<form name="form1" method="post" action="/Profile/MyTask/EditMyTask1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_TaskType"></span>&nbsp;<span class="red">*</span></td>
    <td width="90%" height="20"><%=rs("TASK_TAG")%></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_TaskName"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><input name="taskname" type="text" id="taskname" value="<%=replace(rs("TASK_NAME"),"^","'")%>" size="50"></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="10%" height="20"><span id="inner_MailRecievers"></span></td>
    <td height="20"><table border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
      <tr>
        <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_AvailableRecievers"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_SelectedRecievers"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="4"><select name="fromitem" size="6" multiple id="fromitem">
        <%FactoryRight "U."
		if isnull(rs("RECIEVERS")) or rs("RECIEVERS")="" then%>
		<%= getEngineer("OPTION",""," where USER_CODE is not null "&factorywhereoutsideand," order by U.USER_NAME","") %>
		<%else%>
        <%= getEngineer("OPTION",""," where instr('"&rs("RECIEVERS")&"',USER_CODE)=0 "&factorywhereoutsideand," order by U.USER_NAME","") %>
		<%end if%>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
        <td rowspan="4"><select name="toitem" size="6" multiple id="toitem">
		<%FactoryRight "U."%>
        <%= getEngineer("OPTION",""," where instr('"&rs("RECIEVERS")&"',USER_CODE)>0 "&factorywhereoutsideand," order by U.USER_NAME","") %>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem)"></div></td>
        <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem)"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem)"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
      </tr>
    </table></td>
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
    <td height="20"><span id="inner_SchedulePeriod"></span></td>
    <td height="20"><input name="period" type="checkbox" id="period" onChange="isperiod()" value="1" <%if rs("IS_SCHEDULE")="1" then%>checked<%end if%>></td>
    </tr>
   <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_HappenItem"></span></td>
    <td height="20"><select name="happenitem" disabled id="happenitem" onChange="isweek()">
      <option>--Select--</option>
      <option value="hour" <%if rs("SCHEDULE_TYPE")="hour" then%>selected<%end if%>>each hour</option>
	  <option value="day" <%if rs("SCHEDULE_TYPE")="day" then%>selected<%end if%>>each day</option>
      <option value="week" <%if rs("SCHEDULE_TYPE")="week" then%>selected<%end if%>>each week</option>
	  </select>
      <span id="span_hour" style="visibility: hidden; position: absolute"><input name="hour_interval" type="text" id="hour_interval" value="<%=rs("HOUR_INTERVAL")%>" size="2" maxlength="4">
      </span>
	  <span id="span_week" style="visibility: hidden; position: absolute">
      <input name="week" type="radio" disabled id="week" value="2" <%if rs("WEEK_DAY")="2" then%>checked<%end if%>>
      <span id="inner_Monday"></span>
      <input name="week" type="radio" disabled id="week" value="3" <%if rs("WEEK_DAY")="3" then%>checked<%end if%>>
	  <span id="inner_Tuesday"></span>
	  <input name="week" type="radio" disabled id="week" value="4" <%if rs("WEEK_DAY")="4" then%>checked<%end if%>>
	  <span id="inner_Wendesday"></span>
	  <input name="week" type="radio" disabled id="week" value="5" <%if rs("WEEK_DAY")="5" then%>checked<%end if%>>
	  <span id="inner_Thursday"></span>
	  <input name="week" type="radio" disabled id="week" value="6" <%if rs("WEEK_DAY")="6" then%>checked<%end if%>>
	  <span id="inner_Friday"></span>
	  <input name="week" type="radio" disabled id="week" value="7" <%if rs("WEEK_DAY")="7" then%>checked<%end if%>>
	  <span id="inner_Saturday"></span>
	  <input name="week" type="radio" disabled id="week" value="1" <%if rs("WEEK_DAY")="1" then%>checked<%end if%>>
	  <span id="inner_Sunday"></span>
	  </span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_HappenTime"></span></td>
    <td height="20"><input name="fromdate" type="text" id="fromdate" value="<%=rs("START_TIME")%>" size="10">
    <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script>
	<input name="happentime1" type="text" disabled id="happentime1" value="<%if not isnull(rs("HAPPEN_TIME")) and rs("HAPPEN_TIME")<>"" then%><%=hour(cdate(rs("HAPPEN_TIME")))%><%end if%>" size="2" onChange="time1()">
      :
      <input name="happentime2" type="text" disabled id="happentime2" value="<%if not isnull(rs("HAPPEN_TIME")) and rs("HAPPEN_TIME")<>"" then%><%=minute(cdate(rs("HAPPEN_TIME")))%><%end if%>" size="2" onChange="time2()"></td>
  </tr>
  
  <tr class="t-c-GrayLight">
    <td height="20" colspan="2"><div align="center">
      <input name="oracle_job_id" type="hidden" id="oracle_job_id" value="<%=rs("ORACLE_JOB_ID")%>">
      <input name="profile_task_id" type="hidden" id="profile_task_id" value="<%=id%>">
      <input name="atonce" type="hidden" id="atonce" value="0">
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
<%end if
rs.close%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->