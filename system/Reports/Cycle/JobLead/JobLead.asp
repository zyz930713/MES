<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Reports/Cycle/CycleCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
fromdate2=request("fromdate2")
todate2=request("todate2")
line=trim(request("line"))
factory=trim(request("factory"))
lead_time=request("lead_time")
this_status=request("this_status")
cutdate=request("cutdate")
cuthour=request("cuthour")
cutminute=request("cutminute")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER asc,J.SHEET_NUMBER asc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""

if fromdate="" and todate="" and fromdate2="" and todate2="" then
fromdate=dateadd("d",-15,date())
todate=date()
end if

if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME)='"&lcase(line)&"'"
end if
if factory<>"" then
where=where&" and J.FACTORY_ID='"&factory&"'"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&" 23:59:59','yyyy-mm-dd hh24:mi:ss')"
end if
if fromdate2<>"" then
where=where&" and J.CLOSE_TIME>=to_date('"&fromdate2&"','yyyy-mm-dd')"
end if
if todate2<>"" then
where=where&" and J.CLOSE_TIME<=to_date('"&todate2&"','yyyy-mm-dd')"
end if
if cutdate="" then
	if lead_time="1" then
	where=where&" and CASE WHEN J.STATUS=0 THEN GET_JOB_WORKDAY(J.START_TIME,SYSDATE) ELSE GET_JOB_WORKDAY(J.START_TIME,J.CLOSE_TIME) END>S.LEAD_TIME and S.LEAD_TIME<>0"
	elseif lead_time="2" then
	where=where&" and CASE WHEN J.STATUS=0 THEN GET_JOB_WORKDAY(J.START_TIME,SYSDATE) ELSE GET_JOB_WORKDAY(J.START_TIME,J.CLOSE_TIME) END=S.LEAD_TIME and S.LEAD_TIME<>0"
	elseif lead_time="0" then
	where=where&" and CASE WHEN J.STATUS=0 THEN GET_JOB_WORKDAY(J.START_TIME,SYSDATE) ELSE GET_JOB_WORKDAY(J.START_TIME,J.CLOSE_TIME) END<=S.LEAD_TIME and S.LEAD_TIME<>0"
	end if
else
	cuttime=cutdate&" "&cuthour&":"&cutminute
	if lead_time="1" then
	where=where&" and GET_JOB_WORKDAY(J.START_TIME,to_date('"&cuttime&"','yyyy-mm-dd hh24:mi:ss'))>S.LEAD_TIME and S.LEAD_TIME<>0"
	elseif lead_time="2" then
	where=where&" and GET_JOB_WORKDAY(J.START_TIME,to_date('"&cuttime&"','yyyy-mm-dd hh24:mi:ss'))=S.LEAD_TIME and S.LEAD_TIME<>0"
	elseif lead_time="0" then
	where=where&" and GET_JOB_WORKDAY(J.START_TIME,to_date('"&cuttime&"','yyyy-mm-dd hh24:mi:ss'))<=S.LEAD_TIME and S.LEAD_TIME<>0"
	end if
end if

if this_status<>"" then
where=where&" and J.STATUS="&this_status
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&fromdate="&fromdate&"&todate="&todate&"&fromdate2="&fromdate2&"&todate2="&todate2&"&lead_time="&lead_time&"&isClosed="&isClosed&"&cutdate="&cutdate&"&cuthour="&cuthour&"&cutminute="&cutminute
FactoryRight "J."
SQL="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.PART_NUMBER_TAG,J.JOB_START_QUANTITY,J.JOB_GOOD_QUANTITY,J.START_TIME,J.CLOSE_TIME,CASE WHEN J.CLOSE_TIME IS NULL THEN GET_JOB_WORKDAY(J.START_TIME,SYSDATE) ELSE GET_JOB_WORKDAY(J.START_TIME,J.CLOSE_TIME) END as CYCLE_TIME,nvl(S.LEAD_TIME,0) as LEAD_TIME from JOB J left join PRODUCT_MODEL S on S.ITEM_NAME= J.PART_NUMBER_TAG where J.JOB_NUMBER is not null "&where&factorywhereoutsideand&order

'response.Write(SQL)

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
<script language=JavaScript src="/Functions/JCheck.js" type=text/javascript></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form name="form1" method="get" action="/Reports/Cycle/JobLead/JobLead.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span><span id="inner_Search"></span></span></td>
  </tr>
  <tr>
    <td height="20" width="80"><span id="inner_SearchJobNumber"></span> </td>
    <td height="20" width="100"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td width="80"><span id="inner_SearchPartNumber"></span></td>
    <td width="100"><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>    
    <td height="20" width="80"><span id="inner_Line"></span></td>
    <td height="20"  width="100"><input name="line" type="text" id="line" value="<%=line%>" size="10"></td>
	<td>&nbsp;</td>
  </tr>
  <tr>
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
    <td width="50"><span id="inner_Status"></span></td>
    <td><select name="this_status" id="this_status">
      <option value="">--All-</option>
      <option value="0" <%if this_status="0" then%>Selected<%end if%>>Opened</option>
      <option value="1" <%if this_status="1" then%>Selected<%end if%>>Closed</option>
    </select>
	</td>
    <td align="left">
		<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">
	</td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobLead_Export.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="td_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="td_Model"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_StartTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_CloseTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Yield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_SearchCycleTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Diff"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>" target="_blank"><% =rs("JOB_NUMBER")%>-<% =repeatstring(rs("SHEET_NUMBER"),"0",3)%></a></div></td>
		<td height="20"><div align="center"><% =rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
		<td><div align="center"><% =formatdate(rs("START_TIME"),application("longdateformat"))%></div></td>
		<td><div align="center">
		  <% =formatdate(rs("CLOSE_TIME"),application("longdateformat"))%>&nbsp;
	    </div></td>
		<td><div align="center">
		  <% =formatpercent(csng(rs("JOB_GOOD_QUANTITY"))/csng(rs("JOB_START_QUANTITY")),2,-1)%>
		</div></td>
		<td <%if csng(rs("CYCLE_TIME"))>csng(rs("LEAD_TIME")) then%>class="red"<%end if%>><div align="center"><%=formatnumber(round(csng(rs("CYCLE_TIME"))/60,2),2,-1)%>&nbsp;h</div></td>
		<td <%if csng(rs("CYCLE_TIME"))>csng(rs("LEAD_TIME")) then%>class="red"<%end if%>><div align="center"><%=formatnumber(round(csng(rs("CYCLE_TIME"))/24/60,2),2,-1)%>&nbsp;d</div></td>
	    <td><div align="center"><%=formatnumber(round(csng(rs("LEAD_TIME"))/24/60,2),2,-1)%>&nbsp;d</div></td>
	    <td><div align="center"><%=formatnumber(round((csng(rs("CYCLE_TIME"))-csng(rs("LEAD_TIME")))/24/60,2),2,-1)%>&nbsp;d</div></td>
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
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->