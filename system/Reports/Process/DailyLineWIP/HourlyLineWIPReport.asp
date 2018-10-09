<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=5000%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<%
WIP_ID=request("WIP_ID")
WIP_NAME=trim(request("WIP_NAME"))
line=trim(request("line"))
station=trim(request("station"))
factory_id=trim(request("factory_id"))
from_day=trim(request("from_day"))
to_day=trim(request("to_day"))
ordername=request("ordername")
ordertype=request("ordertype")
if line<>"" then
where=where&" and HWD.LINE_ID='"&line&"'"
end if
if station<>"" then
where=where&" and HWD.STATION_ID='"&station&"'"
end if
if ordername="" and ordertype="" then
order=" order by L.LINE_NAME,S.WIP_SEQUENCY"
else
order=" order by "&ordername&" "&ordertype
end if
if from_day<>""and to_day<>"" then
SQL="Select S.STATION_NAME,L.LINE_NAME,nvl(avg(HWD.QUANTITY),0) as QUANTITY,HWL.REPORT_TIME from BAR_REPORT.HOURLY_WIP_DETAIL HWD inner join BAR_REPORT.HOURLY_WIP_LIST HWL on HWD.WIP_ID=HWL.NID inner join BAR_WEB.LINE L on HWD.LINE_ID=L.NID inner join BAR_WEB.STATION S on HWD.STATION_ID=S.NID where and to_char(HWL.REPORT_TIME,'yyyy-mm-dd')>='"&formatdate(from_day,application("F_shortdateformat"))&"' and to_char(HWL.REPORT_TIME,'yyyy-mm-dd')<='"&formatdate(to_day,application("F_shortdateformat"))&"' and S.WIP_REPORT_COLUMN=1 and S.FACTORY_ID='"&factory_id&"' and QUANTITY<>0 "&where&order
else
SQL="Select S.STATION_NAME,L.LINE_NAME,HWD.QUANTITY,HWL.REPORT_TIME from BAR_REPORT.HOURLY_WIP_DETAIL HWD inner join BAR_REPORT.HOURLY_WIP_LIST HWL on HWD.WIP_ID=HWL.NID inner join BAR_WEB.LINE L on HWD.LINE_ID=L.NID inner join BAR_WEB.STATION S on HWD.STATION_ID=S.NID where HWD.WIP_ID='"&WIP_ID&"' and S.WIP_REPORT_COLUMN=1 and S.FACTORY_ID='"&factory_id&"' and QUANTITY<>0 "&where&order
end if
session("SQL")=SQL
'response.Write(SQL)
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<form action="/Reports/Process/DailyLineWIP/HourlyLineWIPReport.asp" method="get" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="4" class="t-c-greenCopy">Browse Recorded WIP -- <%=WIP_NAME%> (<% =formatdate(from_day,application("longdateformat"))%> to <% =formatdate(to_day,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="4" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('WIPReport_Export.asp?WIP_ID=<%=WIP_ID%>&WIP_NAME=<%=WIP_NAME%>&factory_id=<%=FACTORY_ID%>&from_day=<%=from_day%>&to_day=<%=to_day%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="4">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="20" class="t-b-blue">Line</td>
      <td height="20"><select name="line" id="line">
        <option value="">-- Select --</option>
        <%=getLine("OPTION",line,factorywhereoutside," order by L.LINE_NAME","")%>
      </select></td>
      <td height="20" class="t-b-blue">Station</td>
      <td height="20"><select name="station" id="station">
        <option value="">-- Select --</option>
        <%=getStation("","OPTION",station,factorywhereoutside," order by S.STATION_NAME","","")%>
      </select></td>
      <td height="20"><input name="WIP_ID" type="hidden" id="WIP_ID" value="<%=WIP_ID%>">
        <input name="factory_id" type="hidden" id="factory_id" value="<%=factory_id%>">
        <input name="from_day" type="hidden" id="from_day" value="<%=from_day%>">
        <input name="to_day" type="hidden" id="to_day" value="<%=to_day%>">
        <input name="Select" type="submit" class="t-b-Yellow" id="Select" value="Select"></td>
    </tr>
  </table>  </td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Family<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Station</div></td>
  <td class="t-t-Borrow"><div align="center">WIP Quantity </div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td><div align="center"><%= rs("STATION_NAME") %></div></td>
	<td><div align="center"><%=rs("QUANTITY")%></div></td>
	</tr>
<%
i=i+1
total_station_quantity=total_station_quantity+csng(rs("QUANTITY"))
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <td>&nbsp;</td>
  <td><div align="center"><%=total_station_quantity%></div></td>
  </tr>
<%else%>
  <tr>
    <td height="20" colspan="4"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->