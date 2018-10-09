<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/Functions/GetCalculatedPartWIPZZZ.asp" -->
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
order=" order by J.PART_NUMBER_TAG asc"
else
order=" order by "&ordername&" "&ordertype
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=todate('"&todate&"','yyyy-mm-dd')"
end if
pagename="/Reports/Process/PartWIP/GeneratePartWIP.asp"
SQL="truncate table PART_WIP_DETAIL_TEMP"
rs.open SQL,conn,3,3
FactoryRight "P."
SQL="select distinct 1,J.PART_NUMBER_TAG from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID inner join FACTORY F on P.FACTORY_ID=F.NID where J.JOB_NUMBER not like '%E' and J.STATUS=0 and P.SECTION_ID='"&section&"' "&factorywhereoutsideand&order
session("SQL")=SQL
rs.open SQL,conn,1,3
FactoryRight "S."
SQLT="select 1,S.NID,S.STATION_NAME,S.STATION_CHINESE_NAME,S.WIP_INCLUDED_STATIONS from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID where S.WIP_REPORT_COLUMN=1 and S.SECTION_ID='"&section&"' "&factorywhereoutsideand&" order by S.WIP_SEQUENCY"
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
<script language="JavaScript" src="/Reports/Process/PartWIP/smtPage.js" type="text/javascript"></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Part WIP</td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
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
part_job_numbers=""
slim_part_job_numbers=""
	if not rsT.eof then
		redim quantity(rsT.recordcount-1)
		k=0
		While not rsT.eof
		job_numbers=""			
		quantity(k)=getCalulatedPartWIP(rs("PART_NUMBER_TAG"),rsT("NID"),rsT("WIP_INCLUDED_STATIONS"),job_numbers,rnd_key)
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
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:smtPage('<%=rs("PART_NUMBER_TAG")%>','<%=slim_part_job_numbers%>','<%=part_WIP_name%>','<%=part_WIP_report_time%>','<%=section%>')"><%= rs("PART_NUMBER_TAG") %></span></div></td>
    <%
	  if not rsT.eof then
	  k=0
	  While not rsT.eof%>
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
</tr>
      <form name="form1" method="post" action="/Reports/Process/PartWIP/SavePartWIP.asp">
  <tr>
    <td height="20" colspan="<%=Tcount%>"><span class="strongred">Error jobs:</span> <%= thiserror %></td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>">Generating time: <% =formatdate(create_time,application("longdateformat"))%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="<%=Tcount%>">
        Report name: 
        <input name="part_WIP_name" type="text" id="part_WIP_name">
        <input name="rnd_key" type="hidden" id="rnd_key" value="<%=rnd_key%>">
        <input name="section" type="hidden" id="section" value="<%=section%>">
        <input name="Save" type="submit" id="Save" value="Save This Report">    </td>
  </tr> 
  </form>
    <form name="form2" method="post" action="PartWIPDetail.asp" target="_blank">
  	<input name="part_number_tag" type="hidden" id="part_number_tag" value="">
	<input name="job_numbers" type="hidden" id="job_numbers" value="">
	<input name="part_WIP_name" type="hidden" id="part_WIP_name" value="">
	<input name="part_WIP_report_time" type="hidden" id="part_WIP_report_time" value="">
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