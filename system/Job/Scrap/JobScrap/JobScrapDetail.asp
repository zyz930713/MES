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

pagename="/Reports/Scrap/JobScrap/JobScrapDetail.asp"
FactoryRight "P."
SQL="select J.* from JOB_MASTER_SCRAP J where J.JOB_NUMBER='"&jobnumber&"'"
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
  <td height="20" colspan="7" class="t-c-greenCopy">Browse  Scrap Records </td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy">User: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobScrapDetail_Export.asp?jobnumber=<%=jobnumber%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="7">&nbsp;</td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part</div></td>
  <td class="t-t-Borrow"><div align="center">Line Name</div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Quantity </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Scrap Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Form Code </div></td>
  </tr>
<%
i=1
scrap_quantity=0
if not rs.eof then
while not rs.eof
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
		<td><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
		<td><div align="center"><%= rs("LINE_NAME")%></div></td>
		<td><div align="center"><%= rs("SCRAP_QUANTITY")%></div></td>
		<td><div align="center"><% =formatdate(rs("SCRAP_TIME"),application("longdateformat"))%></div></td>
		<td><div align="center"><%=rs("PROFILE_FORM_ID")%>&nbsp;</div></td>
  </tr>
<%
scrap_quantity=scrap_quantity+cint(rs("SCRAP_QUANTITY"))
i=i+1
rs.movenext
wend
%>
	<tr>
	  <td height="20">&nbsp;</td>
	  <td height="20"><div align="center">Total</div></td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	  <td><div align="center"><%=scrap_quantity%>&nbsp;</div></td>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
  </tr>
<%
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->