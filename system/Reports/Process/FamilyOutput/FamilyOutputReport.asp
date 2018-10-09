<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
family_output_id=request.QueryString("family_output_id")
family_output_name=trim(request("family_output_name"))
report_time=trim(request("report_time"))
factory_id=trim(request.QueryString("factory_id"))
section_id=trim(request.QueryString("section_id"))
from_time=request("from_time")
to_time=request("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by L.LINE_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/FamilyOutput/FamilyOutputRporte.asp"
family_name=getSeriesGroup("TEXT",null," where S.FACTORY_ID='"&factory_id&"'","",",")
station_name=getStation(false,"ENGLISH_TEXT",null," where S.OUTPUT_REPORT_COLUMN=1 and S.SECTION_ID='"&section_id&"' and S.FACTORY_ID='"&factory_id&"'"," order by S.OUTPUT_SEQUENCY",null,",")
if family_name<>"" then
family_name=left(family_name,len(family_name)-1)
a_family_name=split(family_name,",")
end if
if station_name<>"" then
a_station_name=split(station_name,",")
end if
dim total_station_quantity()
redim total_station_quantity(ubound(a_station_name))
Tcount=ubound(a_station_name)+5
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Recorded Output -- <%=family_output_name%> from <%=from_time%> to <%=to_time%> (<% =formatdate(report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('OutputReport_Export.asp?output_id=<%=output_id%>&output_name=<%=output_name%>&output_report_time=<%=output_report_time%>&section_id=<%=section_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Family</div></td>
  <%for i=0 to ubound(a_station_name)%>
  <td class="t-t-Borrow"><div align="center"><%=a_station_name(i)%></div></td>
  <%next%>
  <td class="t-t-Borrow"><div align="center">Total</div></td>
  <td class="t-t-Borrow"><div align="center">Analysis</div></td>
  </tr>
<form action="/Reports/FinalYield/FamilyScrap/FamilyScrapDetail.asp" method="post" name="formDetail" target="_blank">
<%
for i=0 to ubound(a_family_name)
%>
<tr>
  <td height="20"><div align="center"><% =i+1%></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('FamilyOutputLineDetail.asp?family_name=<%=a_family_name(i)%>&fromtime=<%=fromtime%>&totime=<%=totime%>&family_output_name=<%=family_output_name%>&report_time=<%=report_time%>')"><%= a_family_name(i) %></span></div></td>
    <%
	this_family_total=0
	for j=0 to ubound(a_station_name)
		SQL="select OUTPUT from FAMILY_OUTPUT_DETAIL where FAMILY_OUTPUT_ID='"&family_output_id&"' and FAMILY_NAME='"&a_family_name(i)&"' and STATION_NAME='"&a_station_name(j)&"'"
		'response.Write(SQL)
		rs.open SQL,conn,1,3
		if not rs.eof then
		output=rs("OUTPUT")
		else
		output=0
		end if
		rs.close
		this_family_total=this_family_total+csng(output)
		total_station_quantity(j)=total_station_quantity(j)+csng(output)
	  %>
	  <td><div align="center"><%=output%>&nbsp;K</div></td>
	<%next%>
	<td><div align="center"><% =this_family_total%>&nbsp;K</div></td>
	<td><div align="center"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.seriesgroup.value='<%=a_family_name(i)%>';document.formDetail.action='http://<%=application("ChartServer")%>/KES1_Barcode/FamilyProductionHourLine.asp';document.formDetail.submit()"><img src="/Images/IconReportLine.gif"></span></div></td>
  </tr>
<%
next
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <%for i=0 to ubound(total_station_quantity)%>
  <td><div align="center">
      <% =total_station_quantity(i)%>
&nbsp;K</div></td>
  <%next%>
  <td><div align="center">&nbsp;</div></td>
  <td><div align="center"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.seriesgroup.value='Overall';document.formDetail.action='http://<%=application("ChartServer")%>/KES1_Barcode/FamilyProductionHourLine.asp';document.formDetail.submit()"><img src="/Images/IconReportLine.gif"></span></div></td>
</tr>
<input name="factory_id" type="hidden" value="<%=factory_id%>">
<input name="from_time" type="hidden" value="<%=from_time%>">
<input name="to_time" type="hidden" value="<%=to_time%>">
<input name="seriesgroup" type="hidden" value="">
</form>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->