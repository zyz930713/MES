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
SQL="select SG.*,F.FACTORY_NAME,S.SECTION_NAME from SERIES_NEW SG inner join FACTORY F on SG.FACTORY_ID=F.NID left join SECTION S on SG.SECTION_ID=S.NID where SG.FAMILY_ID='"& NID &"' order by SG.SERIES_NAME "
'RESPONSE.WRITE SQL
'RESPONSE.END
rs.open SQL,conn,1,3

%>
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
    <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Series"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_Factory"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_Section"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_InternalYield"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_TargetYield"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime"></span></div></td>
  </tr>
 <%
while not rs.eof 
%>
    <tr>
      <td height="20"><div align="center"><%= rs("SERIES_NAME") %></div></td>
      <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
      <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
      <td><div align="center"><%= formatpercent(csng(rs("TARGET_INTERNALYIELD")),2,-1)%></div></td>
      <td><div align="center"><%= rs("TARGET_YIELD") %>%</div></td>      
      <td><div align="center">
        <%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
        <%=newtime%></div></td>   
    </tr>
    <%
rs.movenext
wend
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->