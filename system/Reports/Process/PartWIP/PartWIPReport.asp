<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/OOCF_Open.asp" -->
<!--#include virtual="/Functions/GetRecordedPartWIP.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
part_WIP_id=request.QueryString("part_WIP_id")
part_WIP_name=trim(request("part_WIP_name"))
part_WIP_report_time=trim(request("part_WIP_report_time"))
section_id=trim(request.QueryString("section_id"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.PART_NUMBER_TAG asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/PartWIP/PartWIPRporte.asp"
SQL="truncate table PART_WIP_DETAIL_TEMP"
'rs.open SQL,conn,3,3
SQL="select distinct 1,J.PART_NUMBER_TAG from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID inner join FACTORY F on P.FACTORY_ID=F.NID where J.JOB_NUMBER not like '%E' and J.STATUS=0 and P.SECTION_ID='"&section_id&"' "&factorywhereoutsideand&order
session("SQL")=SQL
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where WIP_REPORT_COLUMN=1 and SECTION_ID='"&section_id&"' order by WIP_SEQUENCY"
rsT.open SQLT,conn,1,3
Tcount=rsT.recordcount+2
dim total_station_quantity
redim total_station_quantity(rsT.recordcount-1)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Reports/Process/PartWIP/smtPage.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Recorded Part WIP -- <%=part_WIP_name%> (<% =formatdate(part_WIP_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('PartWIPReport_Export.asp?part_WIP_id=<%=part_WIP_id%>&part_WIP_name=<%=part_WIP_name%>&part_WIP_report_time=<%=part_WIP_report_time%>&section_id=<%=section_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <%
  if not rsT.eof then
  While not rsT.eof%>
  <td class="t-t-Borrow"><div align="center"><%=rsT("STATION_NAME")%><br><%=rsT("STATION_CHINESE_NAME")%></div></td>
  <%
  rsT.movenext
  wend
  rsT.movefirst
  end if%>
  </tr>
<%
i=1
dim quantity()
if not rs.eof then
while not rs.eof
part_job_numbers=""
	if not rsT.eof then
		redim quantity(rsT.recordcount-1)
		k=0
		While not rsT.eof
		job_numbers=""
		quantity(k)=getRecordedPartWIP(part_WIP_id,rs("PART_NUMBER_TAG"),rsT("NID"),job_numbers)
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
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:smtPage('<%=rs("PART_NUMBER_TAG")%>','<%=slim_part_job_numbers%>','<%=part_WIP_name%>','<%=part_WIP_report_time%>','<%=section_id%>')"><%= rs("PART_NUMBER_TAG") %></span></div></td>
    <%	if not rsT.eof then
			k=0
			While not rsT.eof
	  %>
	  <div align="center"><td><div align="center"><%=quantity(k)%>&nbsp;</div></td></div>
	<%		
			total_station_quantity(k)=total_station_quantity(k)+csng(quantity(k))
			k=k+1
			rsT.movenext
			wend
		rsT.movefirst
		end if
	%>
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
</tr>
</form>
    <form name="form2" method="post" action="PartWIPDetail.asp" target="_blank">
	<input name="part_number_tag" type="hidden" id="part_number_tag" value="">
	<input name="job_numbers" type="hidden" id="job_numbers" value="">
	<input name="part_WIP_name" type="hidden" id="part_WIP_name" value="">
	<input name="part_WIP_report_time" type="hidden" id="part_WIP_report_time" value="">
	<input name="section_id" type="hidden" id="section_id" value="">
  </form>
<%else%>
  <tr>
    <td height="20" colspan="3"><div align="center">No Records </div></td>
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
<!--#include virtual="/WOCF/OOCF_Close.asp" -->