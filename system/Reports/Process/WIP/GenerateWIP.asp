<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOLE_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/Functions/GetCalculatedWIP.asp" -->
<%
create_time=now()
rnd_key=gen_key(10)
thiserror=""

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
section=trim(request("section"))
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by L.LINE_NAME asc"
else
order=" order by "&ordername&" "&ordertype
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=todate('"&todate&"','yyyy-mm-dd')"
end if
pagename="/Reports/Process/WIP/GenerateWIP.asp"
SQL="truncate table WIP_DETAIL_TEMP"
rs.open SQL,conn,3,3
FactoryRight "L."
SQL="select 1,L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID where L.STATUS=1 and L.SECTION_ID='"&section&"' "&factorywhereoutsideand&order
session("SQL")=SQL
rs.open SQL,conn,1,3
FactoryRight "S."
SQLT="select 1,S.NID,S.STATION_NAME,S.STATION_CHINESE_NAME,S.WIP_INCLUDED_STATIONS from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID where S.WIP_REPORT_COLUMN=1 and S.SECTION_ID='"&section&"' "&factorywhereoutsideand&" order by S.WIP_SEQUENCY"
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
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="JavaScript" src="/Reports/Process/WIP/smtPage.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Line WIP</td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right"></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=L.LINE_NAME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Line <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=L.LINE_NAME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <%
  if not rsT.eof then
  While not rsT.eof%>
    <td class="t-t-Borrow"><div align="center"><%=rsT("STATION_NAME")%><br>
            <%=rsT("STATION_CHINESE_NAME")%></div></td>
    <%
  rsT.movenext
  wend
  rsT.movefirst
  end if%>
    <td class="t-t-Borrow"><div align="center">Total</div></td>
  </tr>
  <%
i=1
total_column_station_quantity=0
if not rs.eof then
while not rs.eof
part_job_numbers=""
slim_part_job_numbers=""
	if not rsT.eof then
		redim quantity(rsT.recordcount-1)
		k=0
		While not rsT.eof
		column_station_quantity=0
		job_numbers=""			
		quantity(k)=getCalulatedWIP(rs("NID"),rs("LINE_NAME"),rsT("NID"),rsT("WIP_INCLUDED_STATIONS"),job_numbers,rnd_key)
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
		for j=0 to ubound(a_part_job_numbers)
			if instr(slim_part_job_numbers,a_part_job_numbers(j))=0 then
			slim_part_job_numbers=slim_part_job_numbers&a_part_job_numbers(j)&","
			end if
		next
		slim_part_job_numbers=left(slim_part_job_numbers,len(slim_part_job_numbers)-1)
	end if
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:smtPage('<%=rs("NID")%>','<%=rs("LINE_NAME")%>','<%=slim_part_job_numbers%>','<%=WIP_name%>','<%=WIP_report_time%>','<%=section%>')"><%= rs("LINE_NAME") %></span></div></td>
    <%
	  if not rsT.eof then
	  k=0
	  While not rsT.eof%>
      <td><div align="center"><%=quantity(k)%>&nbsp;</div></td>
    <%
	  column_station_quantity=column_station_quantity+csng(quantity(k))
	  total_column_station_quantity=total_column_station_quantity+csng(quantity(k))
	  total_station_quantity(k)=total_station_quantity(k)+csng(quantity(k))
	  k=k+1
	  rsT.movenext
	  wend
	  rsT.movefirst
	  end if
	  %>
	  <td><div align="center"><%=column_station_quantity%>&nbsp;</div></td>
  </tr>
  <%
rs.movenext
i=i+1
wend
%>
  <tr class="t-c-GrayLight">
    <td height="20">&nbsp;</td>
    <td><div align="center">Total</div></td>
    <%for k=0 to ubound(total_station_quantity)%>
    <td><div align="center">
      <% =total_station_quantity(k)%>
    </div></td>
    <%next%>
	<td><div align="center">
      <% =total_column_station_quantity%>
    </div></td>
  </tr>
  <form name="form1" method="post" action="/Reports/Process/WIP/SaveWIP.asp">
    <tr>
      <td height="20" colspan="<%=Tcount%>"><span class="strongred">Error jobs:</span> <%= thiserror %></td>
    </tr>
    <tr>
      <td height="20" colspan="<%=Tcount%>">Generating time:
        <% =formatdate(create_time,application("longdateformat"))%>
        &nbsp;</td>
    </tr>
    <tr>
      <td height="20" colspan="<%=Tcount%>"> Report name:
        <input name="WIP_name" type="text" id="WIP_name">
          <input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
          <input name="section" type="hidden" id="section" value="<%=section%>">
          <input name="Save" type="submit" id="Save" value="Save This Report">
      </td>
    </tr>
  </form>
  <form name="form2" method="post" action="WIPLineDetail.asp" target="_blank">
    <input name="line_id" type="hidden" id="line_id" value="">
    <input name="line_name" type="hidden" id="line_name" value="">
    <input name="job_numbers" type="hidden" id="job_numbers" value="">
    <input name="WIP_name" type="hidden" id="WIP_name" value="">
    <input name="WIP_report_time" type="hidden" id="WIP_report_time" value="">
    <input name="section_id" type="hidden" id="section_id" value="">
  </form>
  <%
else
%>
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