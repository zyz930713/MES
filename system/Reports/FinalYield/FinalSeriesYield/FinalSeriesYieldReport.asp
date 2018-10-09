<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
finalseries_id=request.QueryString("finalseries_id")
finalseries_name=trim(request.QueryString("finalseries_name"))
finalseries_report_time=trim(request.QueryString("finalseries_report_time"))
from_time=trim(request.QueryString("from_time"))
to_time=trim(request.QueryString("to_time"))
factory_id=trim(request.QueryString("factory_id"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by Y.SERIES_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/FinalYield/FinalSeriesYield/FinalSeriesYieldReport.asp"
SQL="select FACTORY_TARGET_YIELD,FACTORY_TARGET_FIRSTYIELD from FINAL_SERIESYIELD_LIST where NID='"&finalseries_id&"'"
session("overall_SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_yield=rs("FACTORY_TARGET_YIELD")
factory_target_firstyield=rs("FACTORY_TARGET_FIRSTYIELD")
else
factory_target_yield="0"
factory_target_firstyield="0"
end if
rs.close
SQL="select Y.* from FINAL_SERIESYIELD_DETAIL Y where Y.FINAL_SERIESYIELD_ID='"&finalseries_id&"' and Y.SERIES_NAME<>'OVERALL' "&order
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
<script language="JavaScript" src="/Reports/FinalYield/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy">Browse Recorded Final Series Yields-- <%=finalseries_name%> from <%=from_time%> to <%=to_time%> (<% =formatdate(finalseries_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FinalSeriesYieldReport_Export.asp?finalseries_id=<%=finalseries_id%>&finalseries_name=<%=finalseries_name%>&finalseries_report_time=<%=finalseries_report_time%>&from_time=<%=from_time%>&to_time=<%=to_time%>&factory_id=<%=factory_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="13">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Series Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Failure</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Assembly Input</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Assembly Output</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Stocked Output </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">First Passed  Yield </div></td>
  <td class="t-b-newyear"><div align="center">First Target Yield </div></td>
  <td class="t-t-Borrow"><div align="center">Delta</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
  <td class="t-b-newyear"><div align="center">Internal Yield</div></td>
  <td height="20" class="t-b-newyear"><div align="center">Target Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Delta</div></td>
</tr>
<%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalSeriesYield/FinalSeriesDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_assembly_input=0
total_assembly_output=0
total_final_output=0
while not rs.eof
%>
<tr <%if rs("OVERALL_EXCEPTION")="1" then%>class="today"<%end if%>>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.store_nids.value='<%=rs("STORE_NIDS")%>';document.all.series_name.value='<%=rs("SERIES_NAME")%>';document.formDetail.submit()"><%= rs("SERIES_NAME") %></span></td>
    <td><div align="center">
      <%if csng(rs("ASSEMBLY_INPUT_QUANTITY"))>0 then%>
      <span class="d_link" style="cursor:hand" onClick="javascript:window.open('http://<%=application("ChartServer")%>/KES1_Barcode/FamilyLineLostPareto.asp?factory_id=<%=factory_id%>&seriesgroup=<%=rs("SERIES_NAME")%>&from_time=<%=from_time%>&to_time=<%=to_time%>','_blank')"><img src="/Images/IconReportPareto.gif" width="22" height="22" align="absmiddle"></span>
      <%else%>
      &nbsp;
      <%end if%>
    </div></td>
    <td height="20"><div align="center"><%= rs("ASSEMBLY_INPUT_QUANTITY") %></div></td>
    <td height="20"><div align="center"><%= rs("ASSEMBLY_OUTPUT_QUANTITY") %></div></td>
    <td height="20"><div align="center"><%=rs("OUTPUT_QUANTITY")%></div></td>
    <td height="20" class="<%=getYieldColor(rs("ASSEMBLY_YIELD"),rs("SERIES_TARGET_FIRSTYIELD"))%>"><div align="center"><%= formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1) %></div></td>
    <td><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_FIRSTYIELD"))/100,2,-1)%></div></td>
    <td class="<%=getYieldColor(rs("FINAL_YIELD"),rs("SERIES_TARGET_YIELD"))%>"><div align="center">
      <div align="center">
        <%if csng(rs("ASSEMBLY_YIELD"))<>0 then
	delta=csng(rs("ASSEMBLY_YIELD"))-csng(rs("SERIES_TARGET_FIRSTYIELD"))/100
	else
	delta=0
	end if%>
        <%=formatpercent(delta,2,-1)%> </div>
    </div></td>
    <td height="20" class="<%=getYieldColor(rs("FINAL_YIELD"),rs("SERIES_TARGET_YIELD"))%>"><div align="center"><%=formatpercent(csng(rs("FINAL_YIELD")),2,-1)%>&nbsp;</div></td>
	<td><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_FIRSTYIELD")),2,-1)%></div></td>
	<td height="20"><div align="center"><%=formatpercent(csng(rs("SERIES_TARGET_YIELD"))/100,2,-1)%></div></td>
    <td><div align="center">
        <%if csng(rs("FINAL_YIELD"))<>0 then
	delta=csng(rs("FINAL_YIELD"))-csng(rs("SERIES_TARGET_YIELD"))/100
	else
	delta=0
	end if%>
      <%=formatpercent(delta,2,-1)%></div></td>
</tr>
<%
	if rs("OVERALL_EXCEPTION")="0" then
	total_assembly_input=total_assembly_input+csng(rs("ASSEMBLY_INPUT_QUANTITY"))
	total_assembly_output=total_assembly_output+csng(rs("ASSEMBLY_OUTPUT_QUANTITY"))
	total_final_output=total_final_output+csng(rs("OUTPUT_QUANTITY"))
	end if
rs.movenext
i=i+1
wend
if total_assembly_input<>0 then
total_assembly_yield=total_assembly_output/total_assembly_input
total_final_yield=total_final_output/total_assembly_input
else
total_assembly_yield=0
total_final_yield=0
end if
%>
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td>&nbsp;</td>
    <td height="20"><div align="center"><%=total_assembly_input%></div></td>
    <td height="20"><div align="center"><%=total_assembly_output%></div></td>
    <td height="20"><div align="center"><%=total_final_output%></div></td>
    <td height="20" class="<%=getYieldColor(total_assembly_yield,factory_target_firstyield)%>"><div align="center"><%=formatpercent(total_assembly_yield,2,-1)%>&nbsp;</div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(factory_target_firstyield)/100,2,-1)%></div></td>
    <td class="<%=getYieldColor(total_final_yield,factory_target_yield)%>">&nbsp;</td>
    <td height="20" class="<%=getYieldColor(total_final_yield,factory_target_yield)%>"><%=formatpercent(total_final_yield,2,-1)%></div></td>
    <td>&nbsp;</td>
    <td height="20"><div align="center"><%=formatpercent(csng(factory_target_yield)/100,2,-1)%></div></td>
    <td><div align="center">
      <%if csng(total_final_yield)<>0 then
	delta=csng(total_final_yield)-csng(factory_target_yield)/100
	else
	delta=0
	end if%>
      <%=formatpercent(delta,2,-1)%></div></td>
  </tr>
<input name="store_nids" type="hidden" value="">
<input name="series_name" type="hidden" value="">
</form>
  <tr>
    <td height="20" colspan="13"><table cellspacing="0" cellpadding="0">
      <tr height="18">
        <td width="78" height="18" class="t-b-Green">&nbsp;</td>
        <td colspan="7" width="350">means cum yield meet target</td>
      </tr>
      <tr height="18">
        <td height="18" class="t-b-Yellow"></td>
        <td colspan="7">means delta is between    0% to 0.5%</td>
      </tr>
      <tr height="21">
        <td height="21" class="t-b-Red"></td>
        <td colspan="7">means delta is bigger    than 0.5%</td>
      </tr>
    </table></td>
  </tr>
    <form action="/Reports/FinalYield/FinalSeriesYield/UpdateFinalSeriesYield.asp" method="post" name="form1" target="_self" onSubmit="return formcheckUpdate()">
  <tr>
    <td height="20" colspan="13">Chart Report:
      <input name="isWW" type="checkbox" id="isWW" value="1" <%if wwnumber<>"" then%>checked<%end if%>>
Save this report as WW
<input name="wwnumber" type="text" id="wwnumber" onChange="weeknumbercheck(this)" value="<%=wwnumber%>" size="2">
of
<input name="yearnumber" type="text" id="yearnumber" onChange="yearnumbercheck(this)" value="<%=yearnumber%>" size="4">
year
<input name="monthnumber" type="text" id="monthnumber" onChange="monthnumbercheck(this)" value="<%=monthnumber%>" size="2">
month.
<input name="finalseries_id" type="hidden" id="finalseries_id" value="<%=finalseries_id%>">
<input name="Update" type="submit" id="Update" value="Update"></td>
  </tr>
  </form>
<%else%>
  <tr>
    <td height="20" colspan="13"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->