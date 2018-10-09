<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
familylinelost_id=request.QueryString("familylinelost_id")
familylinelost_name=trim(request.QueryString("familylinelost_name"))
familylinelost_report_time=trim(request.QueryString("familylinelost_report_time"))
from_time=request.QueryString("from_time")
to_time=trim(request.QueryString("to_time"))
factory_id=trim(request.QueryString("factory_id"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by Y.FAMILY_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Process/FamilyLineLost/FamilyLineLostReport.asp"
SQL="select FACTORY_TARGET_QUANTITY,FACTORY_TARGET_AMOUNT,CHART_WEEK,CHART_YEAR,CHART_MONTH from FAMILY_LINELOST_LIST where NID='"&familylinelost_id&"'"
session("overall_SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_quantity=rs("FACTORY_TARGET_QUANTITY")
factory_target_amount=rs("FACTORY_TARGET_AMOUNT")
wwnumber=rs("CHART_WEEK")
yearnumber=rs("CHART_YEAR")
monthnumber=rs("CHART_MONTH")
else
factory_target_quantity="0"
factory_target_amount="0"
end if
rs.close
SQL="select Y.* from FAMILY_LINELOST_DETAIL Y where Y.FAMILY_LINELOST_ID='"&familylinelost_id&"' and FAMILY_NAME<>'OVERALL' "&order
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Reports/FinalYield/FamilyLineLost/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Recorded Family Line Lost -- <%=familylinelost_name%> from <%=formatdate(from_time,application("longdateformat"))%> to <%=formatdate(to_time,application("longdateformat"))%> (<% =formatdate(familylinelost_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FamilyLineLostReport_Export.asp?familylinelost_id=<%=familylinelost_id%>&familylinelost_name=<%=familylinelost_name%>&familylinelost_report_time=<%=familylinelost_report_time%>&from_time=<%=from_time%>&to_time=<%=to_time%>&factory_id=<%=factory_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="9">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Family Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Assembly Input</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Lost Quantity</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Lost Amount </div></td>
  <td height="20" class="t-t-Borrow"><div align="center"> Lost  Percetage </div></td>
  <td class="t-t-Borrow"><div align="center">Target Quantity</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Target Amount </div></td>
  <td class="t-t-Borrow"><div align="center">Analysis</div></td>
</tr>
<%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FamilyLineLost/FamilyLineLostDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_input=0
total_lost_quantity=0
total_lost_amount=0
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.included_system_items.value='<%=rs("INCLUDED_SYSTEM_ITEMS")%>';document.all.seriesgroup.value='<%=rs("FAMILY_NAME")%>';document.formDetail.action='FamilyLineLostDetail.asp';document.formDetail.submit()"><%= rs("FAMILY_NAME") %></span></td>
    <td height="20"><div align="center"><%=rs("INPUT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("LINELOST_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("LINELOST_AMOUNT")%></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("LINELOST_PERCENTAGE")),2,-1)%></div></td>
    <td height="20"><div align="center"><%=rs("FAMILY_TARGET_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("FAMILY_TARGET_AMOUNT")%></div></td>
    <td><div align="center"><%if rs("INCLUDED_SYSTEM_ITEMS")<>"" and rs("LINELOST_QUANTITY")<>"0" then%><span class="d_link" style="cursor:hand" onClick="javascript:document.all.included_system_items.value='<%=rs("INCLUDED_SYSTEM_ITEMS")%>';document.all.seriesgroup.value='<%=rs("FAMILY_NAME")%>';document.formDetail.action='http://<%=application("ChartServer")%>/KES1_Barcode/FamilyLineLostDetailPie.asp';document.formDetail.submit()"><img src="/Images/IconReportPie.gif" width="22" height="22"></span><%else%>&nbsp;<%end if%></div></td>
</tr>
<%
	total_input=total_input+csng(rs("INPUT_QUANTITY"))
	total_lost_quantity=total_lost_quantity+csng(rs("LINELOST_QUANTITY"))
	total_lost_amount=total_lost_amount+csng(rs("LINELOST_AMOUNT"))
rs.movenext
i=i+1
wend
%>
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td height="20"><div align="center"><%=total_input%></div></td>
    <td height="20"><div align="center"><%=total_lost_quantity%></div></td>
    <td height="20"><div align="center"><%=total_lost_amount%></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td>&nbsp;</td>
  </tr>
<input name="included_system_items" type="hidden" value="">
<input name="factory_id" type="hidden" value="<%=factory_id%>">
<input name="from_time" type="hidden" value="<%=from_time%>">
<input name="to_time" type="hidden" value="<%=to_time%>">
<input name="seriesgroup" type="hidden" value="">
</form>
  <tr>
    <td height="20" colspan="9">&nbsp;</td>
  </tr>
  <form action="/Reports/Process/FamilyLineLost/UpdateFamilyLineLost.asp" method="post" name="form1" target="_self" onSubmit="return formcheckUpdate()">
  <tr>
    <td height="20" colspan="9">Chart Report:
    <input name="isWW" type="checkbox" id="isWW" value="1" <%if wwnumber<>"" then%>checked<%end if%>>
	Save this report as WW
	<input name="wwnumber" type="text" id="wwnumber" onChange="weeknumbercheck()" value="<%=wwnumber%>" size="2">
	of 
	<input name="yearnumber" type="text" id="yearnumber" onChange="yearnumbercheck()" value="<%=yearnumber%>" size="4">
	year
<input name="monthnumber" type="text" id="monthnumber" onChange="monthnumbercheck()" value="<%=monthnumber%>" size="2">
month.
	<input name="familylinelost_id" type="hidden" id="familylinelost_id" value="<%=familylinelost_id%>">
	<input name="Update" type="submit" id="Update" value="Update"></td>
  </tr>
  </form>
<%else%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->