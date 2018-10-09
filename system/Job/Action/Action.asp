<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
fromdate=dateadd("d",-7,date())
todate=dateadd("d",1,date())
pagename="/Job/Action/Action.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="javascript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Job/Action/FormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Job/Action/Lan_Action.asp" -->
</head>

<body onLoad="language();">
<form name="form1" method="post" action="/Job/Action/Action1.asp" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_ReportTitle"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_ReportName"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><input name="report_name" type="text" id="report_name" size="50"></td>
    <td><span id="inner_ReportPeriod"></span>&nbsp;<span class="red">*</span></td>
    <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script>
&nbsp;<span id="inner_To"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
    </script>
&nbsp; </td>
    </tr>
  <tr>
    <td height="20"><span id="inner_Factory"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory" onChange="ChangeFactory()">
      <option value="">-- Select Factory --</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION","",factorywhereinside,"","") %>
    </select></td>
    <td><span id="inner_Type"></span>&nbsp;<span class="red">*</span></td>
    <td><select name="report_type" id="report_type">
	  <option value="">-- Select Type --</option>
      <option value="ACTION_PERCENTAGE">ACTION PERCENTAGE</option>
    </select>    </td>
    </tr>
  <tr>
    <td height="20"><span id="inner_ThisStation"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><select name="this_station" id="this_station" onChange="ChangeStation()">
      <option value="">-- Select Station --</option>
    </select></td>
    <td><span id="inner_ThisAction"></span>&nbsp;<span class="red">*</span></td>
    <td><select name="this_action" id="this_action" onChange="ChangeThisAction()">
      <option value="">-- Select Action --</option>
    </select></td>
    </tr>
  <tr>
    <td height="20"><span id="inner_ReferStation"></span>&nbsp;<span class="red">*</span></td>
    <td height="20"><select name="refer_station" id="refer_station" onChange="ChangeReferStation()">
      <option value="">-- Select Station --</option>
    </select></td>
    <td><span id="inner_ReferAction"></span>&nbsp;<span class="red">*</span></td>
    <td><select name="refer_action" id="refer_action" onChange="ChangeReferAction()">
      <option value="">-- Select Action --</option>
    </select></td>
    </tr>
  <tr>
    <td height="20"><span id="inner_Family"></span></td>
    <td height="20" colspan="3"><select name="family" id="family">
      <option value="">-- Select Family --</option>
    </select></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_MailRecievers"></span>&nbsp;<span class="red">*</span></td>
    <td height="20" colspan="3"><table border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
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
  <tr>
    <td height="20" colspan="4"><div align="center">
      <input name="this_action_name" type="hidden" id="this_action_name">
      <input name="refer_action_name" type="hidden" id="refer_action_name">
	  <input name="refer_station_name" type="hidden" id="refer_station_name">
      <input name="Generate" type="submit" id="Generate" value="Generate">
      &nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->