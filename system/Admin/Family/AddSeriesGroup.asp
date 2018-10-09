<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/FAMILY/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/FAMILY/AddSeriesGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
</tr>
<tr> 
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
	<% =session("User") %></td>
</tr>
<tr>
  <td width="100" height="20"><span id="td_Family"></span><span class="red">*</span> </td>
  <td height="20">
      <div align="left">
        <input name="seriesgroupname" type="text" id="seriesgroupname" size="50">
      </div></td>
    </tr>
<tr>
  <td height="20"><span id="td_Factory"></span><span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value=""></option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION","",factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><span id="td_Section"></span>  <span class="red">*</span></td>
  <td height="20"><select name="section" id="section">
    <option value=""></option>
    <%FactoryRight "S."%>
    <%= getSection("OPTION","",factorywhereoutside,"","") %>
  </select>  </td>
</tr>

<tr>
  <td height="20"><span id="td_LeadTime"></span> </td>
  <td height="20"><input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" size="4">
      <select name="lead_time_unit" id="lead_time_unit">
        <option value=""></option>
        <option value="MM">Minutes</option>
        <option value="HH">Hours</option>
        <option value="DD">Days</option>
    </select></td>
</tr>
<!--
<tr>
  <td height="20">WIP Time </td>
  <td height="20"><input name="wip_time" type="text" id="wip_time" onChange="numbercheck(this)" size="4">
    <select name="wip_time_unit" id="wip_time_unit">
      <option value="">-- Select--</option>
      <option value="MM">Minutes</option>
      <option value="HH">Hours</option>
      <option value="DD">Days</option>
    </select></td>
</tr>
-->
<tr>
  <td height="20"><span id="td_FirstPassYield"></span><span class="red">*</span></td>
  <td height="20"><input name="first_yield" type="text" id="first_yield" value="0">
    %</td>
</tr>
<tr>
  <td height="20"><span id="td_InternalYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="internal_yield" type="text" id="internal_yield" value="0">
    %</td>
</tr>
<!--
<tr>
  <td height="20">Target of Retest Yield <span class="red">*</span></td>
  <td height="20"><input name="inspect_yield" type="text" id="inspect_yield" value="0">
    %</td>
</tr>
-->
<tr>
  <td height="20"><span id="td_TargetYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="final_yield" type="text" id="final_yield" value="0">
    %</td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input name="btnReset" type="reset" id="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->