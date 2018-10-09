<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))

pagename="/Reports/Store/JobStore/JobStoreDetail.asp"
FactoryRight "P."
SQL="select J.* from JOB_MASTER_STORE_PRE J where J.JOB_NUMBER='"&jobnumber&"'"
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
  <td height="20" colspan="11" class="t-c-greenCopy">Browse  Store Records </td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy">User: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobStoreDetail_Export.asp?jobnumber=<%=jobnumber%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="11">&nbsp;</td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part</div></td>
  <td class="t-t-Borrow"><div align="center">Line Name</div></td>
  <td class="t-t-Borrow"><div align="center">Input Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Inspect Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Store Quantity </div></td>
  <td class="t-t-Borrow"><div align="center">Ontime Yield</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Store Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Store Type </div></td>
  <td class="t-t-Borrow"><div align="center">Sub Jobs</div></td>
  </tr>
<%
i=1
input_quantity=0
inspect_quantity=0
store_quantity=0
if not rs.eof then
while not rs.eof
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
		<td><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
		<td><div align="center"><%= rs("LINE_NAME")%></div></td>
		<td><div align="center"><%= rs("INPUT_QUANTITY")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("INSPECT_QUANTITY")%></div></td>
		<td><div align="center"><%= rs("STORE_QUANTITY")%></div></td>
		<td><div align="center"><%= formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
		<td><div align="center"><% =formatdate(rs("STORE_TIME"),application("longdateformat"))%></div></td>
		<td><div align="center"><% =rs("STORE_TYPE")%></div></td>
		<td><div align="center"><%=rs("SUB_JOB_NUMBERS")%>&nbsp;</div></td>
  </tr>
<%
if rs("INPUT_QUANTITY")<>"" then
input_quantity=input_quantity+cint(rs("INPUT_QUANTITY"))
end if
if rs("INSPECT_QUANTITY")<>"" then
inspect_quantity=inspect_quantity+cint(rs("INSPECT_QUANTITY"))
end if
if rs("STORE_QUANTITY")<>"" then
store_quantity=store_quantity+cint(rs("STORE_QUANTITY"))
end if
i=i+1
rs.movenext
wend
%>
	<tr>
	  <td height="20">&nbsp;</td>
	  <td height="20"><div align="center">Total</div></td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	  <td><div align="center"><%=input_quantity%>&nbsp;</div></td>
	  <td><div align="center"><%=inspect_quantity%></div></td>
	  <td><div align="center"><%=store_quantity%>&nbsp;</div></td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
  </tr>
<%
else
%>
  <tr>
    <td height="20" colspan="11"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->