<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
if request.QueryString("from_time")<>"" then
from_time=request.QueryString("from_time")
else
from_time=request.Form("fromdate")&" "&request.Form("fromhour")
end if
if trim(request.QueryString("to_time"))<>"" then
to_time=trim(request.QueryString("to_time"))
else
to_time=request.Form("todate")&" "&request.Form("tohour")
end if
profile_task_id=trim(request("profile_task_id"))
factory_id=trim(request("factory_id"))
pagename="/Reports/Yield/DailyLineYield/DailyLineYieldReport.asp"
dateinterval=datediff("d",from_time,to_time)+1
cells=(dateinterval+1)*4+5
family_names=getSeriesGroup("TEXT",""," where S.FACTORY_ID='"&factory_id&"'"," order by S.SERIES_GROUP_NAME","|")
if family_names<>"" then
family_names=left(family_names,len(family_names)-1)
a_family_names=split(family_names,"|")
else
dim a_family_names(0)
end if
line_names=getLine("TEXT",""," where L.FACTORY_ID='"&factory_id&"'"," order by L.LINE_NAME",",")
if line_names<>"" then
line_names=left(line_names,len(line_names)-1)
a_line_names=split(line_names,",")
end if
SQL="select TARGET_FIRSTYIELD from FACTORY where NID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_first_target=csng(rs("TARGET_FIRSTYIELD"))/100
else
factory_first_target=0
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/FinalYield/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=cells%>" class="t-c-greenCopy">Browse Recorded Daily Line Yield -- from <% =formatdate(from_time,application("longdateformat"))%> to <%=formatdate(to_time,application("longdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="<%=cells%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Regenerate Report" onClick="javascript:window.open('RegenerateDailyLineYieldReport.asp?profile_task_id=<%=profile_task_id%>&factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/IconRegenerate.gif" alt="Regenerate Report" width="22" height="22" align="absmiddle"></span>&nbsp;<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyLineYieldReport_Export.asp?profile_task_id=<%=profile_task_id%>&factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL_Middle.gif" alt="Export to Excel" width="22" height="22" align="absmiddle"></span>&nbsp;<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyLineYieldSource_Export.asp?factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>&page_source=daily')">&nbsp;<img src="/Images/IconReportTable.gif" alt="Export the Source" width="22" height="22" align="absmiddle"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=cells%>">&nbsp;</td>
</tr>
<tr>
  <td rowspan="2" class="t-t-Borrow"><div align="center">NO</div></td>
  <td rowspan="2" class="t-t-Borrow"><div align="center">Family Name </div></td>
  <td rowspan="2" class="t-t-Borrow"><div align="center">Line Name</div></td>
  <%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date %>
  <td height="20" colspan="5" class="<%if weekday(current_date)=1 then%>t-b-Red<%elseif weekday(current_date)=7 then%>t-c-greenCopy<%else%>t-t-Borrow<%end if%>"><div align="center"><% =formatdate(current_date,application("longdateformat"))%> (<%=shortweekdayconvert(weekday(current_date))%>)</div></td>
  <%current_date=dateadd("d",1,current_date)
	loop%>
	<td height="20" colspan="5" class="t-b-midautumn"><div align="center">Overall</div></td>
  </tr>
<tr>
	<%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date %>
  <td class="t-t-Borrow"><div align="center">Input</div></td>
  <td class="t-t-Borrow"><div align="center">Output</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Target</div></td>
  <td class="t-t-Borrow"><div align="center">Gap</div></td>
  <%current_date=dateadd("d",1,current_date)
	loop%>
  <td class="t-t-Borrow"><div align="center">Input</div></td>
  <td class="t-t-Borrow"><div align="center">Output</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Target</div></td>
  <td class="t-t-Borrow"><div align="center">Gap</div></td>
</tr>
<%
overall_input_quantity=0
overall_output_quantity=0
overall_yield=0
set rs1=server.CreateObject("adodb.recordset")
for i=0 to ubound(a_family_names)
	for j=0 to ubound(a_line_names)
	SQL="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and LINE_NAME='"&a_line_names(j)&"' and LINE_DAY>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and LINE_DAY<=to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss')"
	'response.Write(SQL&"<br>")
	rs.open SQL,conn,1,3
	if not rs.eof then
	line_input_quantity=0
	line_output_quantity=0
	line_yield=0
	line_target_yield=0
%>
<tr>
  <td height="20"><div align="center"><%=i+1%></div></td>
    <td><div align="center"><span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyLineYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&line_name=<%=a_line_names(j)%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><%=a_family_names(i)%></span></div></td>
    <td height="20"><div align="center"><%=a_line_names(j)%></div></td>
	<%
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date
	SQL1="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and LINE_NAME='"&a_line_names(j)&"' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
	rs1.open SQL1,conn,1,3
	if not rs1.eof then
	assembly_input_quantity=csng(rs1("ASSEMBLY_INPUT_QUANTITY"))
	assembly_output_quantity=csng(rs1("ASSEMBLY_OUTPUT_QUANTITY"))
	assembly_yield=csng(rs1("ASSEMBLY_YIELD"))
	family_target_firstyield=csng(rs1("FAMILY_TARGET_FIRSTYIELD"))
	else
	assembly_input_quantity=0
	assembly_output_quantity=0
	assembly_yield=0
	family_target_firstyield=0
	end if
	rs1.close
	%>
    <td><div align="center"><%=assembly_input_quantity%></div></td>
    <td><div align="center"><%=assembly_output_quantity%></div></td>
    <td height="20" class="<%=getNormalYieldColor(assembly_yield,family_target_firstyield)%>"><div align="center">
	<%if assembly_yield<>0 then%><span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyLineYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&line_name=<%=a_line_names(j)%>&from_time=<%=current_date%>&to_time=<%=dateadd("d",1,current_date)%>')"><%=formatpercent(assembly_yield,2,-1)%></span>
	<%else%>
	<%=formatpercent(assembly_yield,2,-1)%>
	<%end if%></div></td>
	<td class="t-c-bluryellow"><div align="center"><%=formatpercent(family_target_firstyield,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent((assembly_yield-family_target_firstyield),2,-1)%></div></td>
	<%
	line_input_quantity=line_input_quantity+assembly_input_quantity
	line_output_quantity=line_output_quantity+assembly_output_quantity
	if family_target_firstyield<>0 then
	line_target_yield=family_target_firstyield
	end if
	overall_input_quantity=overall_input_quantity+assembly_input_quantity
	overall_output_quantity=overall_output_quantity+assembly_output_quantity
	current_date=dateadd("d",1,current_date)
	loop
	if line_input_quantity<>0 then
	line_yield=line_output_quantity/line_input_quantity
	end if
	%>
	<td><div align="center"><%=line_input_quantity%></div></td>
    <td><div align="center"><%=line_output_quantity%></div></td>
    <td height="20" class="<%=getNormalYieldColor(line_yield,line_target_yield)%>"><div align="center">
	<%if line_yield<>0 then%>
	<span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyLineYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&line_name=<%=a_line_names(j)%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><%=formatpercent(line_yield,2,-1)%></span>
	<%else%>
	<%=formatpercent(line_yield,2,-1)%>
	<%end if%></div></td>
    <td class="t-c-bluryellow"><div align="center"><%=formatpercent(line_target_yield,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent((line_yield-line_target_yield),2,-1)%></div></td>
</tr>
<%	end if
	rs.close
	next
next
%>  
	<tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td><div align="center">Overall</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
	<%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=0
	do while current_date<=end_date
	SQL1="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='OVERALL' and LINE_NAME='OVERALL' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
	rs1.open SQL1,conn,1,3
	if not rs1.eof then
	total_input_quantity=csng(rs1("ASSEMBLY_INPUT_QUANTITY"))
	total_output_quantity=csng(rs1("ASSEMBLY_OUTPUT_QUANTITY"))
	total_yield=csng(rs1("ASSEMBLY_YIELD"))
	else
	total_input_quantity=0
	total_output_quantity=0
	total_yield=0
	end if
	rs1.close
%>
    <td><div align="center"><%=total_input_quantity%></div></td>
    <td><div align="center"><%=total_output_quantity%></div></td>
    <td height="20" class="<%=getNormalYieldColor(total_yield,factory_first_target)%>"><div align="center"><%=formatpercent(total_yield,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent(factory_first_target,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent((total_yield-factory_first_target),2,-1)%></div></td>
	<%if amount_reader=true then%><%end if%>
	<%
	current_date=dateadd("d",1,current_date)
	i=i+1
	loop
	if overall_input_quantity<>0 then
	overall_yield=overall_output_quantity/overall_input_quantity
	end if
	%>
	<td><div align="center"><%=overall_input_quantity%></div></td>
    <td><div align="center"><%=overall_output_quantity%></div></td>
    <td height="20" class="<%=getNormalYieldColor(overall_yield,factory_first_target)%>"><div align="center"><%=formatpercent(overall_yield,2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(factory_first_target,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent((overall_yield-factory_first_target),2,-1)%></div></td>
	</tr>
<%set rs1=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->