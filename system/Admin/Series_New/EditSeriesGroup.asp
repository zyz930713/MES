<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from SERIES_NEW where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Series_New/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/SERIES_NEW/EditSeriesGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span></td>
</tr>
<tr> 
	<td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
     <% =session("User") %></td>
</tr>
<tr>
  <td width="20%" height="20"><div align="left"><span id="td_Series"></span><span class="red">*</span> </div></td>
    <td >
      <div align="left">
        <input name="seriesgroupname" type="text" id="seriesgroupname" value="<%=rs("SERIES_NAME")%>">
      </div></td>
    </tr>
<tr>
  <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value=""></option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><span id="td_Section"></span> <span class="red">*</span></td>
  <td height="20"><select name="section" id="section">
    <option value=""></option>
    <%FactoryRight "S."%>
    <%= getSection("OPTION",rs("SECTION_ID"),factorywhereoutside,"","") %>
  </select>  </td>
</tr>
 <tr>
  <td height="20"><span id="td_Family"></span><span class="red">*</span></td>
  <td height="20"><select name="family" id="family">
    <option value=""></option>
    <%= getFamily("OPTION",rs("FAMILY_ID"),factorywhereoutside,"","") %>
  </select>  </td>
</tr>
<!--
<tr>
  <td height="20">Excepted from Overall </td>
  <td height="20"><input name="overall_exception" type="checkbox" id="overall_exception" value="1" <%if rs("OVERALL_EXCEPTION")="1" then%>checked<%end if%>></td>
  -->
</tr>
<tr>
  <td height="20"><span id="td_LeadTime"></span></td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
    <input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="lead_time_unit" id="lead_time_unit">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
<!--
<tr>
  <td height="20">WIP Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("WIP_TIME")),newtime)%>
    <input name="wip_time" type="text" id="wip_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="wip_time_unit" id="wip_time_unit">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
-->
<tr>
  <td height="20"><span id="td_FirstPassYield"></span><span class="red">*</span></td>
  <td height="20"><input name="first_yield" type="text" id="first_yield" value="<%=rs("TARGET_FIRSTYIELD")%>">
    %</td>
</tr>
<tr>
  <td height="20"><span id="td_InternalYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="internal_yield" type="text" id="internal_yield" value="<%=csng(rs("TARGET_INTERNALYIELD"))*100%>">
%</td>
</tr>
<!--
<tr>
  <td height="20">Target of Retest Yield <span class="red">*</span></td>
  <td height="20"><input name="inspect_yield" type="text" id="inspect_yield" value="<%=csng(rs("TARGET_INSPECTYIELD"))*100%>">
    %</td>
</tr>
-->
<tr>
  <td height="20"><span id="td_TargetYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="final_yield" type="text" id="final_yield" value="<%=rs("TARGET_YIELD")%>">
    %</td>
</tr>
<tr>
  <td height="20" colspan="2">&nbsp;</td>
</tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input type="reset" name="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
