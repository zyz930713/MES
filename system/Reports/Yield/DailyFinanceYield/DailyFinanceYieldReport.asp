<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
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
pagename="/Reports/Yield/DailyFinanceYield/DailyFinanceYieldReport.asp"
dateinterval=datediff("d",from_time,to_time)+1
cells=(dateinterval+1)*6+2
family_names=getFinanceSeriesGroup("TEXT",""," where S.FACTORY_ID='"&factory_id&"'"," order by S.SERIES_GROUP_NAME","|")
if family_names<>"" then
family_names=left(family_names,len(family_names)-1)
a_family_names=split(family_names,"|")
end if
SQL="select FINANCE_TARGET_YIELD,FINANCE_PLAN_TARGET_YIELD from FACTORY where NID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_finance_target=csng(rs("FINANCE_TARGET_YIELD"))
factory_finance_plan_target=csng(rs("FINANCE_PLAN_TARGET_YIELD"))
else
factory_finance_target=0
factory_finance_plan_target=0
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
  <td height="20" colspan="<%=cells%>" class="t-c-greenCopy">Browse Recorded Daily Finance Yield -- from <%=formatdate(from_time,application("longdateformat"))%> to <%=formatdate(to_time,application("longdateformat"))%></td>
</tr>
<tr>
  <td height="20" colspan="<%=cells%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Regenerate Report" onClick="javascript:window.open('RegenerateDailyFinanceYieldReport.asp?profile_task_id=<%=profile_task_id%>&factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/IconRegenerate.gif" alt="Regenerate Report" width="22" height="22" align="absmiddle"></span>&nbsp;<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyFinanceYieldReport_Export.asp?profile_task_id=<%=profile_task_id%>&factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL_Middle.gif" alt="Export to Excel" width="22" height="22" align="absmiddle"></span>&nbsp;<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyFinanceYieldSource_Export.asp?factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')">&nbsp;<img src="/Images/IconReportTable.gif" alt="Export the Source" width="22" height="22" align="absmiddle"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=cells%>">&nbsp;</td>
</tr>
<tr>
  <td rowspan="2" class="t-t-Borrow"><div align="center">NO</div></td>
  <td rowspan="2" class="t-t-Borrow"><div align="center">Family Name </div></td>
  <%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date %>
  <td height="20" colspan="6" class="<%if weekday(current_date)=1 then%>t-b-Red<%elseif weekday(current_date)=7 then%>t-c-greenCopy<%else%>t-t-Borrow<%end if%>"><div align="center"><%=current_date%> (<%=shortweekdayconvert(weekday(current_date))%>)</div></td>
  <%current_date=dateadd("d",1,current_date)
	loop%>
	<td height="20" colspan="6" class="t-b-midautumn"><div align="center">Overall</div></td>
  </tr>
<tr>
	<%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date %>
  <td class="t-t-Borrow"><div align="center">Input Value </div></td>
  <td class="t-t-Borrow"><div align="center">Output Value </div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Value </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Target</div></td>
  <td class="t-t-Borrow"><div align="center">Plan Target </div></td>
  <%current_date=dateadd("d",1,current_date)
	loop%>
  <td class="t-t-Borrow"><div align="center">Input Value </div></td>
  <td class="t-t-Borrow"><div align="center">Output Value </div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Value </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Target</div></td>
  <td class="t-t-Borrow"><div align="center">Plan Target</div></td>
</tr>
<%
overall_input_quantity=0
overall_input_amount=0
overall_output_quantity=0
overall_output_amount=0
overall_scrap_quantity=0
overall_scrap_amount=0
overall_amount_yield=0
for i=0 to ubound(a_family_names)
	line_input_quantity=0
	line_input_amount=0
	line_output_quantity=0
	line_output_amount=0
	line_scrap_quantity=0
	line_scrap_amount=0
	line_yield=0
	line_target_yield=0
	line_plan_target_yield=0
%>
<tr>
  <td height="20"><div align="center"><%=i+1%></div></td>
    <td><div align="center"><span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyFinanceYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><%=a_family_names(i)%></span></div></td>
    <%
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	do while current_date<=end_date
	SQL1="select * from DAILY_FINANCE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
	rs.open SQL1,conn,1,3
	if not rs.eof then
	input_quantity=csng(rs("INPUT_QUANTITY"))
	input_amount=csng(rs("INPUT_AMOUNT"))
	output_quantity=csng(rs("OUTPUT_QUANTITY"))
	output_amount=csng(rs("OUTPUT_AMOUNT"))
	scrap_quantity=csng(rs("SCRAP_QUANTITY"))
	scrap_amount=csng(rs("SCRAP_AMOUNT"))
	amount_yield=csng(rs("AMOUNT_YIELD"))
	family_target_yield=csng(rs("TARGET_YIELD"))
	family_plan_target_yield=csng(rs("PLAN_TARGET_YIELD"))
	else
	input_quantity=0
	input_amount=0
	output_quantity=0
	output_amount=0
	scrap_quantity=0
	scrap_amount=0
	amount_yield=0
	family_target_yield=0
	family_plan_target_yield=0
	end if
	rs.close
	%>
    <td><div align="center"><%=input_amount%></div></td>
    <td><div align="center"><%=output_amount%></div></td>
    <td><div align="center"><%=scrap_amount%></div></td>
    <td height="20" class="<%=getNormalYieldColor(amount_yield,family_target_yield)%>"><div align="center">
	<%if amount_yield<>0 then%><span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyFinanceYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&from_time=<%=current_date%>&to_time=<%=dateadd("d",1,current_date)%>')"><%=formatpercent(amount_yield,2,-1)%></span>
	<%else%>
	<%=formatpercent(amount_yield,2,-1)%>
	<%end if%></div></td>
	<td class="t-c-bluryellow"><div align="center"><%=formatpercent(family_target_yield,2,-1)%></div></td>
	<td class="t-c-bluryellow"><div align="center"><%=formatpercent(family_plan_target_yield,2,-1)%></div></td>
	<%
	line_input_quantity=line_input_quantity+input_quantity
	line_input_amount=line_input_amount+input_amount
	line_output_quantity=line_output_quantity+output_quantity
	line_output_amount=line_output_amount+output_amount
	line_scrap_quantity=line_scrap_quantity+scrap_quantity
	line_scrap_amount=line_scrap_amount+scrap_amount
	if family_target_yield<>0 then
	line_target_yield=family_target_yield
	end if
	if family_plan_target_yield<>0 then
	line_plan_target_yield=family_plan_target_yield
	end if
	overall_input_quantity=overall_input_quantity+input_quantity
	overall_input_amount=overall_input_amount+input_amount
	overall_output_quantity=overall_output_quantity+output_quantity
	overall_output_amount=overall_output_amount+output_amount
	overall_scrap_quantity=overall_scrap_quantity+scrap_quantity
	overall_scrap_amount=overall_scrap_amount+scrap_amount
	current_date=dateadd("d",1,current_date)
	loop
	if line_input_amount<>0 then
	line_yield=line_output_amount/line_input_amount
	end if
	%>
	<td><div align="center"><%=line_input_amount%></div></td>
	<td><div align="center"><%=line_output_amount%></div></td>
	<td><div align="center"><%=line_scrap_amount%></div></td>
    <td height="20" class="<%=getNormalYieldColor(line_yield,line_target_yield)%>"><div align="center">
	<%if line_yield<>0 then%>
	<span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyFinanceYieldDetail.asp?factory_id=<%=factory_id%>&family_name=<%=a_family_names(i)%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><%=formatpercent(line_yield,2,-1)%></span>
	<%else%>
	<%=formatpercent(line_yield,2,-1)%>
	<%end if%></div></td>
    <td class="t-c-bluryellow"><div align="center"><%=formatpercent(line_target_yield,2,-1)%></div></td>
    <td class="t-c-bluryellow"><div align="center"><%=formatpercent(line_plan_target_yield,2,-1)%></div></td>
</tr>
<%
next
%>  
	<tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td><div align="center">Overall</div></td>
    <%current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=0
	do while current_date<=end_date
	SQL1="select * from DAILY_FINANCE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='OVERALL' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
	rs.open SQL1,conn,1,3
	if not rs.eof then
	total_input_quantity=csng(rs("INPUT_QUANTITY"))
	total_input_amount=csng(rs("INPUT_AMOUNT"))
	total_output_quantity=csng(rs("OUTPUT_QUANTITY"))
	total_output_amount=csng(rs("OUTPUT_AMOUNT"))
	total_scrap_quantity=csng(rs("SCRAP_QUANTITY"))
	total_scrap_amount=csng(rs("SCRAP_AMOUNT"))
	total_amount_yield=csng(rs("AMOUNT_YIELD"))
	else
	total_input_quantity=0
	total_input_amount=0
	total_output_quantity=0
	total_output_amount=0
	total_scrap_quantity=0
	total_scrap_amount=0
	total_amount_yield=0
	end if
	rs.close
%>
    <td><div align="center"><%=total_input_amount%></div></td>
    <td><div align="center"><%=total_output_amount%></div></td>
    <td><div align="center"><%=total_scrap_amount%></div></td>
    <td height="20" class="<%=getNormalYieldColor(total_amount_yield,factory_target)%>"><div align="center"><%=formatpercent(total_amount_yield,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent(factory_finance_target,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent(factory_finance_plan_target,2,-1)%></div></td>
	<%
	current_date=dateadd("d",1,current_date)
	i=i+1
	loop
	if overall_input_amount<>0 then
	overall_amount_yield=overall_output_amount/overall_input_amount
	end if
	%>
	<td><div align="center"><%=overall_input_amount%></div></td>
	<td><div align="center"><%=overall_output_amount%></div></td>
	<td><div align="center"><%=overall_scrap_amount%></div></td>
    <td height="20" class="<%=getNormalYieldColor(overall_amount_yield,factory_target)%>"><div align="center"><%=formatpercent(overall_amount_yield,2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(factory_finance_target,2,-1)%></div></td>
	<td><div align="center"><%=formatpercent(factory_finance_plan_target,2,-1)%></div></td>
	</tr>
	<tr>
	  <td height="20" colspan="<%=cells%>"><span class="d_link" style="cursor:hand" onClick="javascript:window.open('DailyFinanceYieldExclude.asp?factory_id=<%=factory_id%>&from_time=<%=from_time%>&to_time=<%=to_time%>')">Open list of excluded scrapping</span></td>
  </tr>
<%set rs=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->