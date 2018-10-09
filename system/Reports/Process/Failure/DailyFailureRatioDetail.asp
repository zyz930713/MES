<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
factory_id=request.QueryString("factory_id")
station_id=request.QueryString("station_id")
defectcode_id=request.QueryString("defectcode_id")
report_day=request.QueryString("report_day")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER,SHEET_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
SQL="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.PART_NUMBER_TAG,J.LINE_NAME,S.STATION_NAME,D.DEFECT_NAME,J.JOB_START_QUANTITY,JS.DEFECT_QUANTITY FROM JOB_DEFECTCODES JS inner join JOB J on JS.JOB_NUMBER=J.JOB_NUMBER and JS.SHEET_NUMBER=J.SHEET_NUMBER inner join STATION S on JS.STATION_ID=S.NID inner join DEFECTCODE D on JS.DEFECT_CODE_ID=D.NID WHERE to_char(J.CLOSE_TIME,'yyyy-mm-dd')='"&formatdate(report_day,application("F_shortdateformat"))&"' and JS.STATION_ID='"&station_id&"'"&order
'response.Write(SQL)
session("SQL")=SQL
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
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Recorded Failure Ratio on <%=report_day%></td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyFailureRatioDetail_Export.asp?station_id=<%=station_id%>&defectcode_id=<%=defectcode_id%>&report_day=<%=report_day%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<!--<tr>
  <td height="20" colspan="9">Sort by: 
    <input name="ByJob" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetailByJobNumber.asp?','_self')">
    Job Number
    <input name="ByPart" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetailByPartNumber.asp?','_self')">
    Part Number </td>
</tr>-->
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Line Name </div></td>
  <td class="t-t-Borrow"><div align="center">Station Name </div></td>
  <td class="t-t-Borrow"><div align="center">Defectcode Name </div></td>
  <td class="t-t-Borrow"><div align="center">Input Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Output Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Ratio</div></td>
  </tr>
<%
i=1
total_input=0
total_output=0
if not rs.eof then
while not rs.eof
%>
<form name="checkform" method="post" action="/Job/SubJobs/BunchJobStationUpdate.asp">
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><span style="cursor:hand" class="d_link" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>','_blank')"><%=rs("JOB_NUMBER")%>-<%=replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)%></span></div></td>
	 <td><div align="center"><%= rs("PART_NUMBER_TAG") %>&nbsp;</div></td>
	 <td><div align="center"><%= rs("LINE_NAME") %></div></td>
	 <td><div align="center"><%= rs("STATION_NAME") %></div></td>
	 <td><div align="center"><%= rs("DEFECT_NAME") %></div></td>
	 <td><div align="center"><%= rs("JOB_START_QUANTITY") %></div></td>
	 <td><div align="center"><%= rs("DEFECT_QUANTITY") %></div></td>
	 <td><div align="center"><%= formatpercent(csng(rs("DEFECT_QUANTITY"))/csng(rs("JOB_START_QUANTITY")),2,-1)%></div></td>
    </tr>
<%
total_input=total_input+csng(rs("JOB_START_QUANTITY"))
total_output=total_output+csng(rs("DEFECT_QUANTITY"))
i=i+1
rs.movenext
wend
if total_input<>0 then
total_yield=total_output/total_input
else
total_yield=0
end if
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <%if DBA=true then%>
  <%end if%>
  <td><div align="center">Total</div></td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td><div align="center"><% =total_input%></div></td>
  <td><div align="center"><%= total_output %></div></td>
  <td><div align="center"><%= formatpercent(total_yield,2,-1) %></div></td>
  </tr>
</form>
<%
else
%>
<tr>
    <td height="20" colspan="9"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->