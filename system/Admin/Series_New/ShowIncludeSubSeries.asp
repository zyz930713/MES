<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
NID=request.querystring("id")
SQL="select SR.*,F.FACTORY_NAME,S.SECTION_NAME,L.LINE_NAME from SUBSERIES SR inner join FACTORY F on SR.FACTORY_ID=F.NID left join SECTION S on SR.SECTION_ID=S.NID left join LINE L on SR.LINE_ID=L.NID  where SR.SERIES_ID='"& NID &"' order by SR.SUBSERIES_NAME "

rs.open SQL,conn,1,3

%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SubSeries"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LineName"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Factory"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Section"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_InternalYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_TargetYield"></span></div></td>
  <!--
  <td class="t-t-Borrow"><div align="center">Exception</div></td>
  -->
  <td class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
  </tr>
 <%
while not rs.eof 
%>
    <tr>
     <td height="20"><div align="center"><%= rs("SUBSERIES_NAME") %></div></td>
    <td><div align="center"><%= rs("LINE_NAME") %>&nbsp;</div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
    <td><div align="center"><%= formatpercent(csng(rs("TARGET_INTERNALYIELD")),2,-1)%></div></td>
    <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>
	<!--
    <td><div align="center"><%if rs("OVERALL_EXCEPTION")="1" then%><img src="/Images/Yes.gif" width="16" height="16"><%else%>&nbsp;<%end if%></div></td>
	-->
	 <td><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('../SubSeries_new/ShowIncludeModel.asp?id=<%=rs("NID")%>')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    </tr>
    <%
rs.movenext
wend
rs.close
%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->