<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
series_name=request.QueryString("series_name")
factory_id=request.QueryString("factory_id")
line_name=request.QueryString("line_name")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Yield/DailyStockYieldYield/DailyStockYieldDetail_Series.asp"
SQL="select INCLUDED_SYSTEM_ITEMS from SERIES where SERIES_NAME='"&series_name&"' and FACTORY_ID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
included_system_items=rs("INCLUDED_SYSTEM_ITEMS")
end if
rs.close
SQL="select * from JOB_MASTER_STORE where FACTORY_ID='"&factory_id&"' and instr('"&included_system_items&"',PART_NUMBER_TAG)>0 and LINE_NAME='"&line_name&"' and STORE_TIME>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and STORE_TIME<=to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') "&order
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
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Recorded Daily Stock Yield of <%=series_name%> -- <%=line_name%> (<%=from_time%> - <%=to_time%>) </td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyStockYieldDetail_Export.asp?series_name=<%=series_name%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Line Name </div></td>
  <td class="t-t-Borrow"><div align="center">Input Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Stock Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Stock Type</div></td>
  <td class="t-t-Borrow"><div align="center">Close Time </div></td>
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
  <%if DBA=true then%>
    <%end if%>
    <td><div align="center"><span style="cursor:hand" class="d_link" onClick="javascript:window.open('/Job/SubJobs/Job.asp?jobnumber=<%=rs("JOB_NUMBER")%>','_blank')"><%=rs("JOB_NUMBER")%></span>
    </div></td>
	 <td><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
	 <td><div align="center"><%= rs("LINE_NAME")%></div></td>
	 <td><div align="center"><%= rs("INPUT_QUANTITY")%></div></td>
	 <td><div align="center"><%= rs("STORE_QUANTITY")%></div></td>
	 <td><div align="center"><%= formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
	 <td><div align="center"><%= rs("STORE_TYPE")%></div></td>
	 <td><div align="center"><%= rs("STORE_TIME")%></div></td>
  </tr>
<%
total_input=total_input+csng(rs("INPUT_QUANTITY"))
total_output=total_output+csng(rs("STORE_QUANTITY"))
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
  <td><div align="center">&nbsp;</div></td>
  <td><div align="center">&nbsp;</div></td>
  <td><div align="center"><% =total_input%></div></td>
  <td><div align="center"><%= total_output %></div></td>
  <td><div align="center"><%= formatpercent(total_yield,2,-1) %></div></td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
<%if DBA=true then%>
  <tr>
    <td height="20" colspan="9"><div align="center">
      <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input name="CheckAll" type="button" id="CheckAll" onClick="checkall()" value="Check All">
  &nbsp;
  <input name="UncheckAll" type="button" id="UncheckAll" onClick="uncheckall()" value="Uncheck All">
  &nbsp;
  <input name="Update" type="submit" id="Update" value="Update">
  &nbsp;
  <input name="Reset" type="reset" id="Reset" value="Reset">
    </div></td>
  </tr>
  <%end if%>
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
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->