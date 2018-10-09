<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetRecordedWIP.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
finalPart_id=request.QueryString("finalPart_id")
finalPart_name=trim(request.QueryString("finalPart_name"))
finalPart_report_time=trim(request.QueryString("finalPart_report_time"))
from_time=trim(request.QueryString("from_time"))
to_time=trim(request.QueryString("to_time"))
factory_id=trim(request.QueryString("factory_id"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by Y.PART_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/FinalYield/FinalPartYield/FinalPartYieldReport.asp"
SQL="select Y.* from FINAL_PARTYIELD_DETAIL Y where Y.FINAL_PartYIELD_ID='"&finalpart_id&"' "&order
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
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Recorded Final Part Yields-- <%=finalPart_name%> from <%=from_time%> to <%=to_time%> (<% =formatdate(finalPart_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FinalPartYieldReport_Export.asp?finalPart_id=<%=finalPart_id%>&finalPart_name=<%=finalPart_name%>&finalPart_report_time=<%=finalPart_report_time%>&from_time=<%=from_time%>&to_time=<%=to_time%>&factory_id=<%=factory_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="9">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Assembly Input</div></td>
  <td class="t-t-Borrow"><div align="center">Assembly Output</div></td>
  <td class="t-t-Borrow"><div align="center">Assembly Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Input Quanity </div></td>
  <td class="t-t-Borrow"><div align="center">Output Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Final Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Target Yield</div></td>
</tr>
<%
i=1
if not rs.eof then
%>
<form action="/Reports/FinalYield/FinalPartYield/FinalPartDetail.asp" method="post" name="formDetail" target="_blank">
<%
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><span class="d_link" style="cursor:hand" onClick="javascript:document.all.store_nids.value='<%=rs("STORE_NIDS")%>';document.all.Part_name.value='<%=rs("PART_NAME")%>';document.formDetail.submit()"><%= rs("PART_NAME") %></span></td>
    <td><div align="center"><%= rs("ASSEMBLY_INPUT_QUANTITY") %></div></td>
    <td><div align="center"><%= rs("ASSEMBLY_OUTPUT_QUANTITY") %></div></td>
    <td><div align="center"><%= formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1) %></div></td>
    <td><div align="center"><%=rs("INPUT_QUANTITY")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("OUTPUT_QUANTITY")%>&nbsp;</div></td>
	<td><div align="center"><%=formatpercent(csng(rs("FINAL_YIELD")),2,-1)%>&nbsp;</div></td>
	    <td><div align="center"><%'=rs("TARGET_YIELD")%>%</div></td>
	  </div>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
<input name="store_nids" type="hidden" value="">
<input name="Part_name" type="hidden" value="">
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