<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<!--#include virtual="/Functions/GetTask.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/Task/NewTask.asp"
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
<!--#include virtual="/Language/Profile/MyTask/Lan_NewMyTask.asp" -->
</head>

<body onLoad="language()">
<div id="paramDiv" style="visibility: hidden; position: absolute"><iframe id="paramFrame"></iframe></div>
<form name="form1" method="post" action="/Profile/MyTask/NewMyTask1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_TaskType"></span>&nbsp;<span class="red">*</span></td>
    <td width="92%" height="20"><select name="task" id="task" onChange="getParam()">
	<option value="">--Select--</option>
	<%= getTask("OPTION","",factorywhereoutsidenull," order by TASK_NAME","") %>
    </select>    </td>
  </tr>
  <tr class="t-c-GrayLight">
    <td width="8%" height="20"><span id="inner_TaskName"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><input name="taskname" type="text" id="taskname" size="50"></td>
  </tr>
    <tr class="t-c-GrayLight">
    <td width="8%" height="20"><span id="inner_MailRecievers"></span></td>
    <td height="20"><table border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
      <tr>
        <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_AvailableRecievers"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_SelectedRecievers"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="4"><select name="fromitem" size="6" multiple id="fromitem">
            <%FactoryRight "U."%>
            <%= getEngineer("OPTION","",factorywhereoutsidenull," order by U.USER_NAME","") %>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem)"></div></td>
        <td rowspan="4"><select name="toitem" size="6" multiple id="toitem">
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
    <td height="20"><input name="period" type="checkbox" id="period" value="1" onChange="isperiod()"></td>
    </tr>
   <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_HappenItem"></span></td>
    <td height="20"><select name="happenitem" disabled id="happenitem" onChange="isweek()">
      <option>--Select--</option>
      <option value="hour">each hour</option>
	  <option value="day">each day</option>
      <option value="week">each week</option>
	  </select>
	  <span id="span_hour" style="visibility: hidden; position: absolute"><input name="hour_interval" type="text" id="hour_interval" value="0" size="2" maxlength="4">
      </span>
	  <span id="span_week" style="visibility: hidden; position: absolute">
      <input name="week" type="radio" disabled id="week" value="2">
      <span id="inner_Monday"></span>
      <input name="week" type="radio" disabled id="week" value="3">
	  <span id="inner_Tuesday"></span>
	  <input name="week" type="radio" disabled id="week" value="4">
	  <span id="inner_Wendesday"></span>
	  <input name="week" type="radio" disabled id="week" value="5">
	  <span id="inner_Thursday"></span>
	  <input name="week" type="radio" disabled id="week" value="6">
	  <span id="inner_Friday"></span>
	  <input name="week" type="radio" disabled id="week" value="7">
	  <span id="inner_Saturday"></span>
	  <input name="week" type="radio" disabled id="week" value="1">
	  <span id="inner_Sunday"></span></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_HappenTime"></span></td>
    <td height="20"><input name="fromdate" type="text" id="fromdate" value="" size="10">
    <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script>
	<input name="happentime1" type="text" disabled id="happentime1" value="<%=hour(now())%>" size="2" onChange="hourcheck(this)">
      :
      <input name="happentime2" type="text" disabled id="happentime2" value="<%=minute(now())%>" size="2" onChange="minutecheck(this)"></td>
  </tr>
  
  <tr class="t-c-GrayLight">
    <td height="20" colspan="2"><div align="center">
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->