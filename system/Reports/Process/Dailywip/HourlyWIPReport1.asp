<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=2000%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetRecordedWIP.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
WIP_ID=request.QueryString("WIP_ID")
WIP_NAME=trim(request.QueryString("WIP_NAME"))
factory_id=trim(request.QueryString("factory_id"))
from_day=trim(request("from_day"))
to_day=trim(request("to_day"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by L.LINE_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/WIP/WIPRporte.asp"
SQL="select L.NID,L.LINE_NAME from LINE L where L.STATUS=1 and L.FACTORY_ID='"&factory_id&"' "&order
session("SQL")=SQL
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where WIP_REPORT_COLUMN=1 and FACTORY_ID='"&factory_id&"' order by WIP_SEQUENCY"
rsT.open SQLT,conn,1,3
Tcount=rsT.recordcount+3
dim total_station_quantity
redim total_station_quantity(rsT.recordcount-1)
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
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Recorded WIP -- <%=WIP_NAME%> (<% =formatdate(from_day,application("longdateformat"))%> to <% =formatdate(to_day,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('WIPReport_Export.asp?WIP_ID=<%=WIP_ID%>&WIP_NAME=<%=WIP_NAME%>&factory_id=<%=FACTORY_ID%>&from_day=<%=from_day%>&to_day=<%=to_day%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Line<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <%
  if not rsT.eof then
  While not rsT.eof%>
  <td class="t-t-Borrow"><div align="center"><%=rsT("STATION_NAME")%><br><%=rsT("STATION_CHINESE_NAME")%></div></td>
  <%
  rsT.movenext
  wend
  rsT.movefirst
  end if%>
  <td class="t-t-Borrow"><div align="center">Total</div></td>
  </tr>
<%
i=1
dim quantity()
if not rs.eof then
while not rs.eof
part_job_numbers=""
total_family_quantity=0
	if not rsT.eof then
		redim quantity(rsT.recordcount-1)
		k=0
		While not rsT.eof
		job_numbers=""
			if from_day<>"" and to_day<>"" then
			quantity(k)=getRecordedWIP_WTD(rs("LINE_NAME"),rsT("NID"),job_numbers,from_day,to_day)
			else
			quantity(k)=getRecordedWIP(WIP_id,rs("LINE_NAME"),rsT("NID"))
			end if
			if job_numbers<>"" then
			part_job_numbers=part_job_numbers&job_numbers&","
			end if
		k=k+1
		rsT.movenext
		wend
	rsT.movefirst
	end if
	if part_job_numbers<>"" then
		part_job_numbers=left(part_job_numbers,len(part_job_numbers)-1)
		a_part_job_numbers=split(part_job_numbers,",")
		slim_part_job_numbers=""
		for j=0 to ubound(a_part_job_numbers)
			if instr(slim_part_job_numbers,a_part_job_numbers(j))=0 then
			slim_part_job_numbers=slim_part_job_numbers&a_part_job_numbers(j)&","
			end if
		next
		slim_part_job_numbers=left(slim_part_job_numbers,len(slim_part_job_numbers)-1)
	end if
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><%= rs("LINE_NAME") %></div></td>
    <%	if not rsT.eof then
			k=0
			While not rsT.eof
	  %>
	  <td><div align="center"><%=quantity(k)%>&nbsp;</div></td>
	<%		
			total_family_quantity=total_family_quantity+csng(quantity(k))
			total_station_quantity(k)=total_station_quantity(k)+csng(quantity(k))
			k=k+1
			rsT.movenext
			wend
		rsT.movefirst
		end if
	%>
	<td><div align="center"><% =total_family_quantity%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%><tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <%for k=0 to ubound(total_station_quantity)%>
  <td><div align="center">
      <% =total_station_quantity(k)%>
&nbsp;</div></td>
  <%next%>
  <td><div align="center">&nbsp;</div></td>
</tr>
</form>
    <form name="form2" method="post" action="../WIP/WIPLineDetail.asp" target="_blank">
  	<input name="line_id" type="hidden" id="line_id" value="">
	<input name="line_name" type="hidden" id="line_name" value="">
	<input name="job_numbers" type="hidden" id="job_numbers" value="">
	<input name="WIP_name" type="hidden" id="WIP_name" value="">
	<input name="WIP_report_time" type="hidden" id="WIP_report_time" value="">
	<input name="section_id" type="hidden" id="section_id" value="">
  </form>
<%else%>
  <tr>
    <td height="20" colspan="4"><div align="center">No Records </div></td>
  </tr>
<%end if
rsT.close
rs.close
set rsT=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->