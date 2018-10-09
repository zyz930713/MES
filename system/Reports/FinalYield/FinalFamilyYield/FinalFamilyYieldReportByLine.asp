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
finalfamily_id=request.QueryString("finalfamily_id")
finalfamily_name=trim(request.QueryString("finalfamily_name"))
finalfamily_report_time=trim(request.QueryString("finalfamily_report_time"))
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
pagename="/Reports/FinalYield/FinalFamilyYield/FinalFamilyYieldReportByLine.asp"
SQL="select FACTORY_TARGET_YIELD,FACTORY_TARGET_FIRSTYIELD,CHART_WEEK,CHART_YEAR,CHART_MONTH from FINAL_FAMILYYIELD_LIST where NID='"&finalfamily_id&"'"
session("overall_SQL")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_yield=rs("FACTORY_TARGET_YIELD")
factory_target_firstyield=rs("FACTORY_TARGET_FIRSTYIELD")
wwnumber=rs("CHART_WEEK")
yearnumber=rs("CHART_YEAR")
monthnumber=rs("CHART_MONTH")
else
factory_target_yield="0"
factory_target_firstyield="0"
end if
rs.close
SQL="select Y.* from FFAMILY_LINEYIELD_DETAIL Y where Y.FINAL_FAMILYYIELD_ID='"&finalfamily_id&"' and FAMILY_NAME<>'OVERALL' "&order
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
  <td height="20" colspan="14" class="t-c-greenCopy">Browse Recorded Final Family Line Yield -- <%=finalfamily_name%> from <%=formatdate(from_time,application("longdateformat"))%> to <%=formatdate(to_time,application("longdateformat"))%> (<% =formatdate(finalfamily_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FinalFamilyYieldReportByLine_Export.asp?finalfamily_id=<%=finalfamily_id%>&finalfamily_name=<%=finalfamily_name%>&finalfamily_report_time=<%=finalfamily_report_time%>&from_time=<%=from_time%>&to_time=<%=to_time%>&factory_id=<%=factory_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Family Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Line Name </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Assembly Input</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Assembly Output</div></td>
  <td class="t-t-Borrow"><div align="center">Retest Output </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Stocked Output </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">First Passed  Yield </div></td>
  <td class="t-b-newyear"><div align="center">First Target Yield </div></td>
  <td class="t-t-Borrow"><p align="center">Retest Yield </p>    </td>
  <td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
  <td class="t-b-newyear"><div align="center">Internal Yield </div></td>
  <td height="20" class="t-b-newyear"><div align="center">Target Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Delta</div></td>
</tr>
<%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalFamilyYield/FinalFamilyDetail.asp" method="post" name="formDetail" target="_blank">
<%
total_assembly_input=0
total_assembly_output=0
total_inspect_output=0
total_final_output=0
while not rs.eof
%>
<tr <%if rs("OVERALL_EXCEPTION")="1" then%>class="today"<%end if%>>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td height="20"><span class="d_link" style="cursor:hand" onClick="javascript:document.all.store_nids.value='<%=rs("STORE_NIDS")%>';document.all.family_name.value='<%=rs("FAMILY_NAME")%>';document.all.line_name.value='<%=rs("LINE_NAME")%>';document.formDetail.submit()"><%= rs("FAMILY_NAME") %></span></td>
    <td height="20"><div align="center"><%= rs("LINE_NAME") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("ASSEMBLY_INPUT_QUANTITY") %></div></td>
    <td height="20"><div align="center"><%= rs("ASSEMBLY_OUTPUT_QUANTITY") %></div></td>
    <td><div align="center"><%=rs("INSPECT_QUANTITY")%></div></td>
    <td height="20"><div align="center"><%=rs("OUTPUT_QUANTITY")%></div></td>
    <td height="20" class="<%=getFirstYieldColor(rs("ASSEMBLY_OUTPUT_QUANTITY"),rs("ASSEMBLY_YIELD"),rs("FAMILY_TARGET_FIRSTYIELD"))%>"><div align="center"><%= formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1) %></div></td>
    <td height="20"><div align="center"><%=formatpercent(csng(rs("FAMILY_TARGET_FIRSTYIELD"))/100,2,-1)%></div></td>
    <td class="<%=getNormalYieldColor(rs("INSPECT_YIELD"),rs("FAMILY_TARGET_INSPECTYIELD"))%>"><div align="center"><%=formatpercent(csng(rs("INSPECT_YIELD")),2,-1)%></div></td>
    <td height="20" class="<%=getYieldColor(rs("FINAL_YIELD"),rs("FAMILY_TARGET_YIELD"))%>"><div align="center"><%=formatpercent(csng(rs("FINAL_YIELD")),2,-1)%></div></td>
	<td><div align="center"><%=formatpercent(csng(rs("FAMILY_TARGET_INTERNALYIELD")),2,-1)%></div></td>
	<td height="20"><div align="center"><%=formatpercent(csng(rs("FAMILY_TARGET_YIELD"))/100,2,-1)%></div></td>
    <td height="20"><div align="center">
        <%if csng(rs("FINAL_YIELD"))<>0 then
	delta=csng(rs("FINAL_YIELD"))-csng(rs("FAMILY_TARGET_YIELD"))/100
	else
	delta=0
	end if%>
        <%=formatpercent(delta,2,-1)%></div></td>
</tr>
<%
	if rs("OVERALL_EXCEPTION")="0" then
	total_assembly_input=total_assembly_input+csng(rs("ASSEMBLY_INPUT_QUANTITY"))
	total_assembly_output=total_assembly_output+csng(rs("ASSEMBLY_OUTPUT_QUANTITY"))
	total_inspect_output=total_inspect_output+csng(rs("INSPECT_QUANTITY"))
	total_final_output=total_final_output+csng(rs("OUTPUT_QUANTITY"))
	end if
rs.movenext
i=i+1
wend
if total_assembly_input<>0 then
total_assembly_yield=total_assembly_output/total_assembly_input
total_inspect_yield=total_inspect_output/total_assembly_input
total_final_yield=total_final_output/total_assembly_input
else
total_assembly_yield=0
total_inspect_yield=0
total_final_yield=0
end if
%>
  <tr class="t-b-Pink">
    <td height="20">&nbsp;</td>
    <td height="20">Overall</td>
    <td><div align="center">&nbsp;</div></td>
    <td height="20"><div align="center"><%=total_assembly_input%></div></td>
    <td height="20"><div align="center"><%=total_assembly_output%></div></td>
    <td><div align="center"><%=total_inspect_output%></div></td>
    <td height="20"><div align="center"><%=total_final_output%></div></td>
    <td height="20" class="<%=getFirstYieldColor(total_assembly_output,total_assembly_yield,factory_target_firstyield)%>"><div align="center"><%=formatpercent(total_assembly_yield,2,-1)%></div></td>
    <td><div align="center"><%=formatpercent(csng(factory_target_firstyield)/100,2,-1)%></div></td>
    <td class="<%=getNormalYieldColor(total_inspect_yield,factory_target_inspectyield)%>"><div align="center"><%=formatpercent(total_inspect_yield,2,-1)%></div></td>
    <td height="20" class="<%=getYieldColor(total_final_yield,factory_target_yield)%>"><div align="center"><%=formatpercent(total_final_yield,2,-1)%></div></td>
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
<input name="family_name" type="hidden" value="">
<input name="line_name" type="hidden" value="">
</form>
  <tr>
    <td height="20" colspan="14"><table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#333333" bordercolordark="#FFFFFF">
      <tr height="18">
        <td width="52" rowspan="3">First Yield </td>
        <td width="50" height="18"><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Green">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td width="201">Means cum yield meet target</td>
      </tr>
      <tr height="18">
        <td height="18"><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Yellow">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means delta is between 0% to 3%</td>
      </tr>
      <tr height="21">
        <td height="21"><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Red">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>If output&gt;10K and delta is bigger than 3ге. If output&lt;10K and  delta is bigger than 5ге.</td>
      </tr>
      <tr height="21">
        <td rowspan="3">Final Yield </td>
        <td><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Green">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means cum yield meet target</td>
      </tr>
      <tr height="21">
        <td><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Yellow">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means delta is between 0% to 0.5%</td>
      </tr>
      <tr height="21">
        <td><table width="50" border="0" cellspacing="0" cellpadding="0">
            <tr class="t-b-Red">
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td>Means delta is bigger than 0.5%</td>
      </tr>
    </table></td>
  </tr>
  <form action="/Reports/FinalYield/FinalFamilyYield/UpdateFinalFamilyYield.asp" method="post" name="form1" target="_self" onSubmit="return formcheckUpdate()">
  <tr>
    <td height="20" colspan="14">Chart Report:
    <input name="isWW" type="checkbox" id="isWW" value="1" <%if wwnumber<>"" then%>checked<%end if%>>
	Save this report as WW
	<input name="wwnumber" type="text" id="wwnumber" onChange="weeknumbercheck()" value="<%=wwnumber%>" size="2">
	of 
	<input name="yearnumber" type="text" id="yearnumber" onChange="yearnumbercheck()" value="<%=yearnumber%>" size="4">
	year
<input name="monthnumber" type="text" id="monthnumber" onChange="monthnumbercheck()" value="<%=monthnumber%>" size="2">
month.
	<input name="finalfamily_id" type="hidden" id="finalfamily_id" value="<%=finalfamily_id%>">
	<input name="Update" type="submit" id="Update" value="Update"></td>
  </tr>
  </form>
<%else%>
  <tr>
    <td height="20" colspan="14"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->