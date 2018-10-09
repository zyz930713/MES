<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
store_nids=request.Form("store_nids")
Part_name=trim(request.Form("Part_name"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/FinalYield/FinalPartYield/FinalPartDetail.asp"
if store_nids<>"" then
store_nids=left(store_nids,len(store_nids)-1)
end if
SQL="select * from JOB_MASTER_STORE where instr('"&store_nids&"',NID)>0 "&order
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
  <td height="20" colspan="8" class="t-c-greenCopy">Browse Recorded Part Yield -- <%=Part_name%></td>
</tr>
<tr>
  <td height="20" colspan="8" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('PartDetail_Export.asp?part_name=<%=part_name%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<form name="formSort" method="post" action="">
<tr>
  <td height="20" colspan="8">Sort by: 
    <input name="ByJob" type="radio" value="1" onClick="document.formSort.action='FinalPartDetailbyJobNumber.asp';document.all.store_nids.value='<%=store_nids%>';document.all.part_name.value='<%=part_name%>';document.formSort.submit()">
    Job Number
    <input name="ByPart" type="radio" value="1" onClick="document.formSort.action='FinalPartDetailbyPartNumber.asp';document.all.store_nids.value='<%=store_nids%>';document.all.part_name.value='<%=part_name%>';document.formSort.submit()">
    Part Number </td>
</tr>
<input name="store_nids" type="hidden" value="">
<input name="part_name" type="hidden" value="">
</form>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Input Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Store Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Yield</div></td>
  <td class="t-t-Borrow"><div align="center">Store Type </div></td>
  <td class="t-t-Borrow"><div align="center">Store Time </div></td>
  </tr>
<%
i=1
total_input=0
total_output=0
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("JOB_NUMBER") %></a></div></td>
	  <div align="center"><td><div align="center"><%= rs("PART_NUMBER_TAG") %>&nbsp;</div></td>
	    <td><div align="center"><%= rs("INPUT_QUANTITY") %></div></td>
	    <td><div align="center"><%= rs("STORE_QUANTITY") %></div></td>
	    <td><div align="center"><%= formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
	    <td><div align="center"><%= rs("STORE_TYPE") %></div></td>
	    <td><div align="center"><%= rs("STORE_TIME") %></div></td>
	  </div>
  </tr>
<%
total_input=total_input+cint(rs("INPUT_QUANTITY"))
total_output=total_output+cint(rs("STORE_QUANTITY"))
i=i+1
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <td><div align="center">&nbsp;</div></td>
	<td><div align="center">
	  <% =total_input%>
    </div></td>
	<td><div align="center"><%= total_output %></div></td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<%
else
%>

  <tr>
    <td height="20" colspan="8"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->