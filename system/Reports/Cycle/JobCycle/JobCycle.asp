<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Cycle/CycleCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/CompareLeadTime.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER asc,J.SHEET_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" then
fromdate=dateadd("d",-1,date())
end if
if todate="" then
todate=date()
end if

if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&" 23:59:59','yyyy-mm-dd hh24:mi:ss')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Cycle/JobCycle/JobCycle.asp"
FactoryRight "P."
SQL="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.PART_NUMBER_TAG,J.START_TIME,J.CLOSE_TIME,J.SHIFT_IN_TIME,J.SHIFT_OUT_TIME,J.CYCLE_TIME from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null"&where&factorywhereoutsideand&order
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language=JavaScript src="/Functions/TablePlus.js" type=text/javascript></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form name="form1" method="post" action="/Reports/Cycle/JobCycle/JobCycle.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr align="center">
    <td height="20" width="80"><span id="inner_SearchJobNumber"></span> </td>
    <td height="20" width="100"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td width="80"><span id="inner_SearchPartNumber"></span></td>
    <td width="100"><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td width="120"><span id="inner_SearchJobStartTime"></span>&nbsp;<span id="inner_SearchFrom"></span></td>
    <td width="100">
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.fromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
      </script>
	</td>
	<td width="30">
		<span id="inner_SearchTo"></span>
    </td>
	<td width="100">
	<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
	<script language=JavaScript type=text/javascript>
		function calendar2Callback(date, month, year)
		{
		document.all.todate.value=year + '-' + month + '-' + date
		}
		calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
	</script>
	</td>
    <td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobCycle_Export.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="td_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="td_Model"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_StartTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_CloseTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")

dim thisjobnumber()
dim thissheetnumber()
dim thisjobtype()
dim thispartnumbertag()
dim thisstart_time()
dim thisclose_time()
dim thiselapsed_time()
redim thisjobnumber(0)
redim thissheetnumber(0)
redim thisjobtype(0)
redim thispartnumbertag(0)
redim thisstart_time(0)
redim thisclose_time(0)
redim thiselapsed_time(0)
current_jobnumber=rs("JOB_NUMBER")
t=1
while not rs.eof and i<=rs.pagesize 
	shift_interval=0
	elapsed_time=0
	if rs("CYCLE_TIME")="" then
		if rs("SHIFT_IN_TIME")<>"" and rs("SHIFT_OUT_TIME")<>"" then
			shift_in_time=left(rs("SHIFT_IN_TIME"),len(rs("SHIFT_IN_TIME"))-1)
			a_shift_in_time=split(shift_in_time,",")
			shift_out_time=left(rs("SHIFT_OUT_TIME"),len(rs("SHIFT_OUT_TIME"))-1)
			a_shift_out_time=split(shift_out_time,",")
			for i=0 to ubound(a_shift_in_time)
			shift_interval=shift_interval+datediff("n",a_shift_out_time(i),a_shift_in_time(i))
			next
		end if
		elapsed_time=cstr(datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))-shift_interval)
		rs("CYCLE_TIME")=elapsed_time
		rs.update
	else
	elapsed_time=datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))
	end if
	
	if rs("JOB_NUMBER")=current_jobnumber then 'Get total parameters' value of all the same jobs
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thisjobtype(UBound(thisjobtype))=rs("JOB_TYPE")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisclose_time(UBound(thisclose_time))=rs("CLOSE_TIME")
	thiselapsed_time(UBound(thiselapsed_time))=elapsed_time
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thisjobtype(UBound(thisjobtype)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisclose_time(UBound(thisclose_time)+1)
	ReDim Preserve thiselapsed_time(UBound(thiselapsed_time)+1)
	else
	%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')"><input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td>&nbsp;</td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
  </tr>
	<tbody id="tab<%=current_jobnumber%>" style="display:none">
	<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
	%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>&jobtype=<%=thisjobtype(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td height="20"><div align="center"><%= thispartnumbertag(u) %>&nbsp;</div></td>
		<td><div align="center"><% =formatdate(thisstart_time(u),application("longdateformat"))%></div></td>
		<td><div align="center">
		  <% =formatdate(thisclose_time(u),application("longdateformat"))%>&nbsp;
	    </div></td>
		<td><div align="center"><%day_time=jobworkday(thisstart_time(u),thisclose_time(u),hourtime)%><%=CompareLeadtime(thispartnumbertag(u),day_time,hourtime)%>&nbsp;d</div></td>
		<td><div align="center"><%if thiselapsed_time(u)<>"" then%><%=formatnumber(thiselapsed_time(u)/60,2,-1)%><%end if%>&nbsp;h</div></td>
	  </tr>
	  
<%		next%>
	</tbody>
<%
	redim thisjobnumber(0)
	redim thissheetnumber(0)
	redim thispartnumbertag(0)
	redim thisstart_time(0)
	redim thisclose_time(0)
	redim thiselapsed_time(0)
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thisjobtype(UBound(thisjobtype))=rs("JOB_TYPE")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisclose_time(UBound(thisclose_time))=rs("CLOSE_TIME")
	thiselapsed_time(UBound(thiselapsed_time))=elapsed_time
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisclose_time(UBound(thisclose_time)+1)
	ReDim Preserve thiselapsed_time(UBound(thiselapsed_time)+1)
	t=t+1
	end if
i=i+1
current_jobnumber=rs("JOB_NUMBER")
rs.movenext
wend
%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')"><input name="tabstatus<%=current_jobnumber%>" type="hidden" value="0"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td>&nbsp;</td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>" style="display:none">
	<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>&jobtype=<%=thisjobtype(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td height="20"><div align="center"><%= thispartnumbertag(u) %>&nbsp;</div></td>
		<td><div align="center"><% =formatdate(thisstart_time(u),application("longdateformat"))%></div></td>
		<td><div align="center">
		  <% =formatdate(thisclose_time(u),application("longdateformat"))%>
	    &nbsp;</div></td>
		<td><div align="center"><%day_time=jobworkday(thisstart_time(u),thisclose_time(u),hourtime)%><%=CompareLeadtime(thispartnumbertag(u),day_time,hourtime)%>&nbsp;d</div></td>
		<td><div align="center"><%if thiselapsed_time(u)<>"" then%><%=formatnumber(thiselapsed_time(u)/60,2,-1)%><%end if%>&nbsp;h</div></td>
    </tr>
<%
		next%>
  </tbody>
<%
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center"><span id="inner_Records"></span></div></td>
  </tr>
<%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->