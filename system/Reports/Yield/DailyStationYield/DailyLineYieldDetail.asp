<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
family_name=request.QueryString("family_name")
factory_id=request.QueryString("factory_id")
line_name=request.QueryString("line_name")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER,SHEET_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
SQL="select INCLUDED_SYSTEM_ITEMS from SERIES_GROUP where SERIES_GROUP_NAME='"&family_name&"' and FACTORY_ID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
included_system_items=rs("INCLUDED_SYSTEM_ITEMS")
end if
rs.close
SQL="select * from JOB where FACTORY_ID='"&factory_id&"' and instr('"&included_system_items&"',PART_NUMBER_TAG)>0 and LINE_NAME='"&line_name&"' and CLOSE_TIME>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and CLOSE_TIME<=to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss') "&order
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
  <td height="20" colspan="8" class="t-c-greenCopy">Browse Recorded Daily Line Yield of <%=family_name%> -- <%=line_name%> (<%=from_time%> - <%=to_time%>) </td>
</tr>
<tr>
  <td height="20" colspan="8" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('DailyLineYieldDetail_Export.asp?family_name=<%=family_name%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<tr>
  <td height="20" colspan="8">Sort by: 
    <input name="ByJob" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetailByJobNumber.asp?family_name=<%=family_name%>&factory_id=<%=factory_id%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>','_self')">
    Job Number
    <input name="ByPart" type="radio" value="1" onClick="javascript:window.open('DailyLineYieldDetailByPartNumber.asp?family_name=<%=family_name%>&factory_id=<%=factory_id%>&line_name=<%=line_name%>&from_time=<%=from_time%>&to_time=<%=to_time%>','_self')">
    Part Number </td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <%if DBA=true then%>
  <%end if%>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Line Name </div></td>
  <td class="t-t-Borrow"><div align="center">Station Name</div></td>
  <td class="t-t-Borrow"><div align="center">Input Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Defectcode Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Yield</div></td>
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
    <td><div align="center"><span style="cursor:hand" class="d_link" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>','_blank')"><%=rs("JOB_NUMBER")%>-<%=replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)%></span>&nbsp;
      <span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('/Job/SubJobs/JobStationUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
    </div></td>
	 <td><div align="center"><%= rs("PART_NUMBER_TAG") %>&nbsp;</div></td>
	 <td><div align="center"><%= rs("LINE_NAME") %></div></td>
	 <td><div align="center"></div></td>
	 <td><div align="center"><%= rs("JOB_START_QUANTITY") %></div></td>
	 <td><div align="center"><%= rs("JOB_GOOD_QUANTITY") %></div></td>
	 <td><div align="center"><%= formatpercent(csng(rs("JOB_ASSEMBLY_YIELD")),2,-1)%></div></td>
    </tr>
<%
total_input=total_input+csng(rs("JOB_START_QUANTITY"))
total_output=total_output+csng(rs("JOB_GOOD_QUANTITY"))
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
  <td>&nbsp;</td>
  <td><div align="center"><% =total_input%></div></td>
  <td><div align="center"><%= total_output %></div></td>
  <td><div align="center"><%= formatpercent(total_yield,2,-1) %></div></td>
  </tr>
<%if DBA=true then%>
  <tr>
    <td height="20" colspan="8"><div align="center">
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
    <td height="20" colspan="8"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->