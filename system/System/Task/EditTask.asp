<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="NewEvent.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from TASK where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/System/Task/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form name="form1" method="post" action="/System/Task/EditTask1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Edit a Task </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">User:
      <% =session("User") %></td>
    </tr>
    <tr> 
      <td width="158" height="20">Task Name <span class="red">*</span></td>
      <td width="683" height="20"><input name="taskname" type="text" id="taskname" value="<%=rs("TASK_NAME")%>" size="50">	  </td>
    </tr>
    <tr>
      <td height="20">Task Chinese Name </td>
      <td height="20"><input name="taskchinesename" type="text" id="taskchinesename" value="<%=rs("TASK_CHINESE_NAME")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="20">Description</td>
      <td height="20"><input name="description" type="text" id="description" value="<%=rs("DESCRIPTION")%>" size="50">      </td>
    </tr>
    <tr>
      <td height="20">Package</td>
      <td height="20"><input name="package" type="text" id="package" value="<%=rs("PACKAGE")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Apply by user </td>
      <td height="20"><input name="apply_by_user" type="checkbox" id="apply_by_user" value="1" <%if rs("APPLY_BY_USER")="1" then%>checked<%end if%>></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 1 </td>
      <td height="20"><select name="paramtype1" id="paramtype1">
        <option value="">--Select Type--</option>
        <option value="text" <%if rs("PARAM_TYPE1")="text" then%>selected<%end if%>>Text</option>
        <option value="radio" <%if rs("PARAM_TYPE1")="radio" then%>selected<%end if%>>Radio</option>
        <option value="option" <%if rs("PARAM_TYPE1")="option" then%>selected<%end if%>>Option</option>
        <option value="checkbox" <%if rs("PARAM_TYPE1")="checkbox" then%>selected<%end if%>>Check</option>
		<option value="calendar" <%if rs("PARAM_TYPE1")="calendar" then%>selected<%end if%>>Calendar</option>
		<option value="period" <%if rs("PARAM_TYPE1")="period" then%>selected<%end if%>>Period</option>
		<option value="time" <%if rs("PARAM_TYPE1")="time" then%>selected<%end if%>>Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 1 </td>
      <td height="20"><input name="param1" type="text" id="param1" value="<%=rs("PARAM1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 1 </td>
      <td height="20"><input name="cparam1" type="text" id="cparam1" value="<%=rs("PARAM_CHINESE1")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 2 </td>
      <td height="20"><select name="paramtype2" id="paramtype2">
        <option value="">--Select Type--</option>
        <option value="text" <%if rs("PARAM_TYPE2")="text" then%>selected<%end if%>>Text</option>
        <option value="radio" <%if rs("PARAM_TYPE2")="radio" then%>selected<%end if%>>Radio</option>
        <option value="option" <%if rs("PARAM_TYPE2")="option" then%>selected<%end if%>>Option</option>
        <option value="checkbox" <%if rs("PARAM_TYPE2")="checkbox" then%>selected<%end if%>>Check</option>
		<option value="calendar" <%if rs("PARAM_TYPE2")="calendar" then%>selected<%end if%>>Calendar</option>
		<option value="period" <%if rs("PARAM_TYPE2")="period" then%>selected<%end if%>>Period</option>
		<option value="time" <%if rs("PARAM_TYPE2")="time" then%>selected<%end if%>>Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 2 </td>
      <td height="20"><input name="param2" type="text" id="param2" value="<%=rs("PARAM2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 2 </td>
      <td height="20"><input name="cparam2" type="text" id="cparam2" value="<%=rs("PARAM_CHINESE2")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 3 </td>
      <td height="20"><select name="paramtype3" id="paramtype3">
        <option value="">--Select Type--</option>
        <option value="text" <%if rs("PARAM_TYPE3")="text" then%>selected<%end if%>>Text</option>
        <option value="radio" <%if rs("PARAM_TYPE3")="radio" then%>selected<%end if%>>Radio</option>
        <option value="option" <%if rs("PARAM_TYPE3")="option" then%>selected<%end if%>>Option</option>
        <option value="checkbox" <%if rs("PARAM_TYPE3")="checkbox" then%>selected<%end if%>>Check</option>
		<option value="calendar" <%if rs("PARAM_TYPE3")="calendar" then%>selected<%end if%>>Calendar</option>
		<option value="period" <%if rs("PARAM_TYPE3")="period" then%>selected<%end if%>>Period</option>
		<option value="time" <%if rs("PARAM_TYPE3")="time" then%>selected<%end if%>>Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 3 </td>
      <td height="20"><input name="param3" type="text" id="param3" value="<%=rs("PARAM3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 3 </td>
      <td height="20"><input name="cparam3" type="text" id="cparam3" value="<%=rs("PARAM_CHINESE3")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Control Type 4 </td>
      <td height="20"><select name="paramtype4" id="paramtype4">
          <option value="">--Select Type--</option>
          <option value="text" <%if rs("PARAM_TYPE4")="text" then%>selected<%end if%>>Text</option>
          <option value="radio" <%if rs("PARAM_TYPE4")="radio" then%>selected<%end if%>>Radio</option>
          <option value="option" <%if rs("PARAM_TYPE4")="option" then%>selected<%end if%>>Option</option>
          <option value="checkbox" <%if rs("PARAM_TYPE4")="checkbox" then%>selected<%end if%>>Check</option>
		  <option value="calendar" <%if rs("PARAM_TYPE4")="calendar" then%>selected<%end if%>>Calendar</option>
		  <option value="period" <%if rs("PARAM_TYPE4")="period" then%>selected<%end if%>>Period</option>
	      <option value="time" <%if rs("PARAM_TYPE4")="time" then%>selected<%end if%>>Time</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Parameter Name 4 </td>
      <td height="20"><input name="param4" type="text" id="param4" value="<%=rs("PARAM4")%>" size="50"></td>
    </tr>
    <tr>
      <td height="20">Parameter Chinese Name 4 </td>
      <td height="20"><input name="cparam4" type="text" id="cparam4" value="<%=rs("PARAM_CHINESE4")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="Submit" value="Submit">
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