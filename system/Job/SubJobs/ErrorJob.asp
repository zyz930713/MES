<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
jobstatus=trim(request("jobstatus"))
line=trim(request("line"))
factory=trim(request("factory"))
timespan=trim(request("timespan"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if jobstatus<>"" then
where=where&" and J.STATUS="&jobstatus
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME)='"&lcase(line)&"'"
end if
if factory<>"" then
where=where&" and P.FACTORY_ID='"&factory&"'"
end if
if timespan<>"" then
where=where&" and round(to_number(JS.close_TIME-JS.start_time)*1440)<="&timespan
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
else
fromdate=dateadd("d",-7,date())
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&sheetnumber="&sheetnumber&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&timespan="&timespan&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Job/ErrorJob.asp"
SQL="select 1, JS.JOB_NUMBER,JS.SHEET_NUMBER,JS.JOB_TYPE,S.STATION_NAME,S.STATION_CHINESE_NAME,JS.OPERATOR_CODE,P.PART_NUMBER,J.PART_NUMBER_TAG,J.LINE_NAME,J.START_TIME,JS.START_TIME,JS.CLOSE_TIME from JOB_STATIONS JS inner join JOB J on JS.JOB_NUMBER=J.JOB_NUMBER and JS.SHEET_NUMBER=J.SHEET_NUMBER and JS.JOB_TYPE=J.JOB_TYPE inner join PART P on J.PART_NUMBER_ID=P.NID inner join STATION S on JS.STATION_ID=S.NID where J.STATUS=1 "&where&order
rs.open SQL,conn,1,3
session("SQL")=SQL
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Job/SubJobs/Lan_ErrorJob.asp" -->
</head>

<body onLoad="language();language_page();language_jobnote()">
<form action="/Job/SubJobs/ErrorJob.asp" method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search"></span></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchJobNumber"></span></td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
      <td><span id="inner_SearchPartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
      <td><span id="inner_SearchStatus"></span></td>
      <td><select name="jobstatus" id="jobstatus">
        <option value="">All</option>
        <option value="0" <%if jobstatus="0" then%>selected<%end if%>>Opened</option>
        <option value="2" <%if jobstatus="2" then%>selected<%end if%>>Paused</option>
        <option value="1" <%if jobstatus="1" then%>selected<%end if%>>Closed</option>
		<option value="3" <%if jobstatus="3" then%>selected<%end if%>>Locked</option>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchLineName"></span></td>
      <td height="20"><input name="line" type="text" id="line" value="<%=line%>"></td>
      <td><span id="inner_SearchFactory"></span></td>
      <td><select name="factory" id="factory">
          <option value="">Factory</option>
          <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
          <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
          <%FactoryRight ""%>
          <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
      </select></td>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchElapsedTime"></span></td>
      <td height="20">&lt;=
        <input name="timespan" type="text" id="timespan" value="<%=timespan%>" size="4">
       <span id="inner_SearchM"></span></td>
      <td><span id="inner_SearchJobStartTime"></span></td>
      <td><span id="inner_SearchFrom"></span>
        <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
      <td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="10"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="t-c-greenCopy">
          <tr>
            <td width="40%"><span id="inner_Browse"></span></td>
            <td width="60%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('ErrorJob_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Station"></span></div></td>
	  <td class="t-t-Borrow"><div align="center"><span id="inner_Operator"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_PartType"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_ElapsedTime"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_StartTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_CloseTime"></span></div></td>
    </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
    <tr>
      <td height="20"><div align="center">
        <% =(session("strpagenum")-1)*pagesize_s+i%>
</div></td>
	  <td><div align="center"><%=rs("STATION_NAME")%>&nbsp;<%=rs("STATION_CHINESE_NAME")%></div></td>
	  <td><div align="center"><%=rs("OPERATOR_CODE")%></div></td>
      <td height="20"><div align="center" class="d_link"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>" target="_blank"><%=rs("JOB_NUMBER")%>-<%=replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)%></a></div></td>
      <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
      <td><div align="center"><%=rs("PART_NUMBER")%></div></td>
      <td><div align="center"><%=rs("LINE_NAME")%></div></td>
      <td><div align="center"><%=datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))%>&nbsp;<%=unit%>&nbsp;m</div></td>
      <td height="20"><div align="center"><% =formatdate(rs("START_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
      <td><div align="center"><% =formatdate(rs("CLOSE_TIME"),application("longdateformat"))%>&nbsp;</div></td>
    </tr>
<%
i=i+1
rs.movenext
wend
else
%>
    <tr>
      <td height="20" colspan="10"><div align="center"><span id="inner_Records"></span></div></td>
    </tr>
<%end if
rs.close%>
	<tr>
      <td height="20" colspan="10"><!--#include virtual="/Components/JobNote.asp" --></td>
    </tr>
</table>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
