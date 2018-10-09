<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from LINE_LABOUR where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Job/Labour/LabourWeekFormCheck.js" type="text/javascript"></script>
<!--#include virtual="/Language/Job/Labour/Lan_AddLabourWeek.asp" -->
<script language="javascript">
function checkhour(ob)
{
	if (!(isNumberString(ob.value,"1234567890"))||new Number(ob.value)>24)
	{
		alert("时间格式错误！");
		ob.value="00";
	}	
	else
	{
		if (ob.value=="24")
		{
		ob.value="00";
		}
	}
}
function checkminute(ob)
{
	if (!(isNumberString(ob.value,"1234567890"))||new Number(ob.value)>60)
	{
		alert("时间格式错误！");
		ob.value="00";
	}	
	else
	{
		if (ob.value=="60")
		{
		ob.value="00";
		}
	}
}
</script>
</head>

<body onLoad="language()">
<form action="/Job/Labour/EditLabourWeek1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20"><span id="inner_HostLine"></span> <span class="red">*</span> </td>
  <td height="20" class="red"><select name="line" id="line">
  <option value="">--Select Line--</option>
  <%'=getLine("OPTION",""," where LEADER='"&session("CODE")&"'","","")%>
  <%=getLine("OPTION_LINE_NAME",rs("LINE_NAME")," where LEADER='3125'","","")%>
  </select>    &nbsp;</td>
</tr>
<tr>
  <td height="20"><span id="inner_Year"></span> <span class="red">*</span></td>
  <td height="20" class="red"><input name="year" type="text" id="year" value="<%=rs("YEAR_INDEX")%>" size="4"></td>
</tr>
<tr>
  <td width="82" height="20"><span id="inner_Week"></span> <span class="red">*</span></td>
    <td width="672" height="20">
      <div align="left">
        <input name="week" type="text" id="week" value="<%=rs("WEEK_INDEX")%>" size="2">
		<span id="inner_From"></span>&nbsp;
        <input name="fromdate" type="text" id="fromdate" value="<%=datevalue(rs("WEEK_FROM_TIME"))%>" size="10" readonly="true">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
        <input name="fromhour" type="text" id="fromhour" value="<%=hour(rs("WEEK_FROM_TIME"))%>" size="2" onChange="checkhour(this)">
 :        
 <input name="fromminute" type="text" id="fromminute" value="<%=minute(rs("WEEK_FROM_TIME"))%>" size="2" onChange="checkminute(this)">
        &nbsp;<span id="inner_To"></span>
<input name="todate" type="text" id="todate" value="<%=datevalue(rs("WEEK_END_TIME"))%>" size="10" readonly="true">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar4Callback');
                        </script>
&nbsp;
<input name="tohour" type="text" id="tohour" value="<%=hour(rs("WEEK_END_TIME"))%>" size="2" onChange="checkhour(this)">
      :
      <input name="tominute" type="text" id="tominute" value="<%=minute(rs("WEEK_END_TIME"))%>" size="2" onChange="checkminute(this)">
      </div></td>
    </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input name="Save" type="submit" id="Save" value="Save">
&nbsp;
<input name="Reset" type="reset" id="Reset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<%rs.close%>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/Functions/TableControl.asp" -->
