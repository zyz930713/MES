<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory_id=request("factory_id")
start_time=request("from_time")
end_time=request("to_time")
jobs=request("jobs")
seriesgroup=request("seriesgroup")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JOB_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/Process/FamilyScrap/FamilyScrapDetail.asp"
SQL="select JS.JOB_NUMBER,JS.SCRAP_QUANTITY,JS.REASON,JS.PART_NUMBER_TAG,JS.LINE_NAME from JOB_MASTER_SCRAP JS where JOB_NUMBER in ('"&replace(jobs,",","','")&"') and JS.SCRAP_TIME>=to_date('"&start_time&"','yyyy-mm-dd hh24:mi:ss') and JS.SCRAP_TIME<=to_date('"&end_time&"','yyyy-mm-dd hh24:mi:ss') order by JS.REASON,JS.JOB_NUMBER"
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
  <td height="20" colspan="7" class="t-c-greenCopy">Browse Family Line Scrap Detail -- <%=seriesgroup%></td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('FinalFamilyDetail_Export.asp?family_name=<%=family_name%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<form name="formSort" method="post" action="">
<!--<tr>
  <td height="20" colspan="7">Sort by: 
    <input name="ByJob" type="radio" value="1" onClick="document.formSort.action='FinalFamilyDetailbyJobNumber.asp';document.all.store_nids.value='<%=store_nids%>';document.all.family_name.value='<%'=family_name%>';document.formSort.submit()">
    Job Number
    <input name="ByPart" type="radio" value="1" onClick="document.formSort.action='FinalFamilyDetailbyPartNumber.asp';document.all.store_nids.value='<%=store_nids%>';document.all.family_name.value='<%'=family_name%>';document.formSort.submit()">
    Part Number </td>
</tr>
--><input name="store_nids" type="hidden" value="">
<input name="family_name" type="hidden" value="">
</form>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Defect Name </div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Quanity
  </div></td>
  <td class="t-t-Borrow"><div align="center">Scrap Amount </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Part Number </div></td>
  <td class="t-t-Borrow"><div align="center">Line Name </div></td>
  </tr>
<%
i=1
total_scrap_quantity=0
total_scrap_amount=0
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><%= rs("REASON") %></div></td>
    <td><div align="center"><%= rs("SCRAP_QUANTITY") %></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><a href="/Job/Job/Job.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><strong><span style="cursor:hand"><%=rs("JOB_NUMBER")%></span></strong></a></div></td>
	  <div align="center"><td><div align="center"><%= rs("PART_NUMBER_TAG") %>&nbsp;</div></td>
	    <td><div align="center"><%= rs("LINE_NAME") %></div></td>
      </div>
  </tr>
<%
total_scrap_quantity=total_scrap_quantity+cint(rs("SCRAP_QUANTITY"))
total_scrap_amount=total_scrap_amount+0
i=i+1
rs.movenext
wend
%>
<tr class="t-c-GrayLight">
  <td height="20">Total</td>
  <td>&nbsp;</td>
  <td><div align="center"><%=total_scrap_quantity%></div></td>
  <td>&nbsp;</td>
  <td><div align="center">&nbsp;</div></td>
  <td><div align="center">&nbsp;</div></td>
	<td><div align="center">&nbsp;</div></td>
  </tr>
<%
else
%>

  <tr>
    <td height="20" colspan="7"><div align="center">No Records</div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->