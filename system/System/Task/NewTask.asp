<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<%
pagename="NewEngineer.asp"
path=request.QueryString("path")
query=request.QueryString("query")
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Task/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form name="form1" method="post" action="/System/Task/NewTask1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Create a New Task </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">User:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="170" height="20">Task Name <span class="red">*</span></td>
      <td width="671" height="20"><input name="taskname" type="text" id="taskname" size="50">	  </td>
    </tr>
    <tr>
      <td height="20"><p>Task Chinese Name <span class="red">*</span></p>      </td>
      <td height="20"><input name="taskchinesename" type="text" id="taskchinesename" size="50"></td>
    </tr>
    <tr> 
      <td height="20">Description</td>
      <td height="20"><input name="description" type="text" id="description" size="50">      </td>
    </tr>
    <tr>
      <td height="20">Apply by user </td>
      <td height="20"><input name="apply_by_user" type="checkbox" id="apply_by_user" value="1"></td>
    </tr>
    <tr>
      <td height="20">Package</td>
      <td height="20"><input name="package" type="text" id="package" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 1 </td>
      <td height="20"><select name="paramtype1" id="paramtype1">
        <option value="">--Select Type--</option>
        <option value="text">Text</option>
        <option value="radio">Radio</option>
        <option value="option">Option</option>
		<option value="checkbox">Check</option>
		<option value="calendar">Calendar</option>
		<option value="period">Period</option>
		<option value="time">Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 1</td>
      <td height="20"><input name="param1" type="text" id="param1" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 1</td>
      <td height="20"><input name="cparam1" type="text" id="cparam1" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 2</td>
      <td height="20"><select name="paramtype2" id="paramtype2">
        <option value="">--Select Type--</option>
        <option value="text">Text</option>
        <option value="radio">Radio</option>
        <option value="option">Option</option>
		<option value="checkbox">Check</option>
		<option value="calendar">Calendar</option>
		<option value="period">Period</option>
		<option value="time">Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 2</td>
      <td height="20"><input name="param2" type="text" id="param2" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 2</td>
      <td height="20"><input name="cparam2" type="text" id="cparam2" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 3 </td>
      <td height="20"><select name="paramtype3" id="paramtype3">
          <option value="">--Select Type--</option>
          <option value="text">Text</option>
          <option value="radio">Radio</option>
          <option value="option">Option</option>
		  <option value="checkbox">Check</option>
		  <option value="calendar">Calendar</option>
		  <option value="period">Period</option>
		  <option value="time">Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 3 </td>
      <td height="20"><input name="param3" type="text" id="param3" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 3 </td>
      <td height="20"><input name="cparam3" type="text" id="cparam3" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 4 </td>
      <td height="20"><select name="paramtype4" id="paramtype4">
          <option value="">--Select Type--</option>
          <option value="text">Text</option>
          <option value="radio">Radio</option>
          <option value="option">Option</option>
          <option value="checkbox">Check</option>
		  <option value="calendar">Calendar</option>
		  <option value="period">Period</option>
		  <option value="time">Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 4 </td>
      <td height="20"><input name="param4" type="text" id="param4" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 4 </td>
      <td height="20"><input name="cparam4" type="text" id="cparam4" size="50"></td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="Submit" value="Submit">
          &nbsp; 
          <input type="reset" name="Submit2" value="Reset">
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->