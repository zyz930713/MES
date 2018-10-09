<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/SeriesGroup/IsCapacity.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from SERIES_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/SeriesGroup/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="deselectedcount()">
<form action="/Admin/SeriesGroup/CapacitySeriesGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Series Group   </td>
</tr>
<tr>
  <td width="153" height="20">Series Group Name</td>
    <td width="601" height="20" class="red"><div align="left"><%=rs("SERIES_GROUP_NAME")%></div></td>
    </tr>
<tr>
  <td height="20">Belonged Factory</td>
  <td height="20"><%=rs("FACTORY_ID")%></td>
</tr>
<tr>
  <td height="20">Section</td>
  <td height="20"><%=rs("SECTION_ID")%></td>
</tr>
<tr>
  <td height="20">Included Parts</td>
  <td height="20"><%=highlightsamestring(rs("INCLUDED_SYSTEM_ITEMS"),",")%>&nbsp;</td>
</tr>
<tr>
  <td height="20">Lead Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%><%=newtime%><%=unit%></td>
</tr>
<tr>
  <td height="20">WIP Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("WIP_TIME")),newtime)%><%=newtime%><%=unit%></td>
</tr>
<tr>
  <td height="20">Target of First-pased Yield</td>
  <td height="20"><%=rs("TARGET_FIRSTYIELD")%> %</td>
</tr>
<tr>
  <td height="20">Target of Internal Yield</td>
  <td height="20"><%=csng(rs("TARGET_INTERNALYIELD"))*100%> %</td>
</tr>
<tr>
  <td height="20">Target of Retest Yield</td>
  <td height="20"><%=csng(rs("TARGET_INSPECTYIELD"))*100%> %</td>
</tr>
<tr>
  <td height="20">Target of Final Yield</td>
  <td height="20"><%=rs("TARGET_YIELD")%> %</td>
</tr>
<tr>
  <td height="20">Capacity</td>
  <td height="20"><input name="capacity" type="text" id="capacity" value="<%=rs("CAPACITY")%>"></td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
