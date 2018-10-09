<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetWIPJobQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
line_name=request.Form("line_name")
job_numbers=request.Form("job_numbers")
WIP_name=trim(request.Form("WIP_name"))
WIP_report_time=trim(request.Form("WIP_report_time"))
section_id=trim(request.Form("section_id"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER asc,J.SHEET_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
pagename="/Reports/WIP/WIPPartDetail.asp"
a_job_numbers=split(job_numbers,",")
jobwhere=""
for i=0 to ubound(a_job_numbers)' each job number with sheet number
	a_this_job_number=split(a_job_numbers(i),"-")
	this_job_number=""
	this_sheet_number=""
	for j=0 to ubound(a_this_job_number)
		if j<>ubound(a_this_job_number) then
		this_job_number=this_job_number&a_this_job_number(j)&"-"
		else
		this_sheet_number=cint(a_this_job_number(j))
		end if
	next
	this_job_number=left(this_job_number,len(this_job_number)-1)
	jobwhere=jobwhere&"(J.JOB_NUMBER='"&this_job_number&"' and J.SHEET_NUMBER='"&this_sheet_number&"') or "
next
jobwhere=left(jobwhere,len(jobwhere)-4)
SQL="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE from JOB J where "&jobwhere&order
session("SQL")=SQL
'response.Write(SQL)
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME,WIP_INCLUDED_STATIONS from STATION where WIP_REPORT_COLUMN=1 and SECTION_ID='"&section_id&"' order by WIP_SEQUENCY"
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
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Recorded WIP -- <%=line_name%> of <%=WIP_name%> (<% =formatdate(WIP_report_time,application("longdateformat"))%>)</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('WIPLineDetail_Export.asp?line_name=<%=line_name%>&WIP_name=<%=WIP_name%>&WIP_report_time=<%=WIP_report_time%>&section_id=<%=section_id%>')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td></tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
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
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =i%></div></td>
    <td><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>" target="_blank"><%= rs("JOB_NUMBER") %>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></a></div></td>
    <%
	  if not rsT.eof then
	  j=0
	  While not rsT.eof%>
	  <div align="center"><td><div align="center"><%station_quantity=getWIPJobQuantity(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("JOB_TYPE"),rsT("NID"),rsT("WIP_INCLUDED_STATIONS"))%><%=station_quantity%>&nbsp;</div></td></div>
	  <%
	  total_station_quantity(j)=total_station_quantity(j)+csng(station_quantity)
	  j=j+1
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
%>
<tr class="t-c-GrayLight">
  <td height="20">&nbsp;</td>
  <td><div align="center">Total</div></td>
  <%for k=0 to ubound(total_station_quantity)%>
  <td><div align="center">
        <% =total_station_quantity(k)%>
        &nbsp;</div></td>
	<%next%>
</tr>
      <form name="form1" method="post" action="/Reports/Process/WIP/SaveWIP.asp">
       </form>
<%
else
%>

  <tr>
    <td height="20" colspan="<%=Tcount%>"><div align="center">No Records</div></td>
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