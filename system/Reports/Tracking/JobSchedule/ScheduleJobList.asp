<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
schedule_id=request("schedule_id")
jobnumber=request("jobnumber")
partnumber=request("partnumber")
formdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by JS.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if schedule_id<>"" then
where=where&" and JS.SCHEDULE_ID='"&schedule_id&"'"
end if
if jobnumber<>"" then
where=where&" and JS.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if fromdate<>"" then
where=where&" and JS.START_TIME>=TO_DATE('"&fromdate&"','YYYY-MM-DD')"
end if
if todate<>"" then
where=where&" and JS.START_TIME<=TO_DATE('"&todate&"','YYYY-MM-DD')"
end if
pagepara="&schedule_id="&schedule_id&"&part_id="&part_id
pagename="/Reports/Tracking/JobSchedule/ScheduleJobList.asp"
SQL="select * from JOB_SCHEDULE_LIST where NID='"&schedule_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
weekindex=rs("WEEK_INDEX")
schedule_name=rs("SCHEDULE_NAME")
end if
rs.close
dim PART_ID(),PART_NAME()
SQL="select NID,PART_NUMBER from PART where STATUS=1"
rs.open SQL,conn,1,3
redim PART_ID(rs.recordcount-1),PART_NAME(rs.recordcount-1)
i=0
while not rs.eof
PART_ID(i)=rs("NID")
PART_NAME(i)=rs("PART_NUMBER")
i=i+1
rs.movenext
wend
rs.close
SQL="select 1,JS.TRACKING_NUMBER,JS.JOB_NUMBER,JS.PART_NUMBER_ID,JS.QUANTITY,P.PART_NUMBER,JS.START_TIME,JS.COMPLETE_TIME from JOB_SCHEDULE_DETAIL JS left join PART P on JS.PART_NUMBER_ID=P.NID where JS.JOB_NUMBER is not null "&where&order
session("SQL")=SQL
rs.open SQL,conn,1,3
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
<script language=javascript src="/Functions/Progress.js" type=text/javascript></script>
</head>

<body>
 <div id="JProgress" style="position:absolute; left:411px; top:244px; z-index:1; visibility: hidden;">
   <table width="200" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#009900" bordercolordark="#FFFFFF">
     <tr>
       <td><table width="350" border="0" cellpadding="0" cellspacing="0">
         <tr>
           <td height="20" class="t-c-greenCopy">
             <div align="center">The progress of transaction</div></td>
         </tr>
         <tr>
           <td height="20" class="t-c-GrayLight"><div><img src="/Images/Progress.gif" id="pprogress" name="pprogress" width="1" height="12"></div></td>
         </tr>
         <tr>
           <td height="20" class="t-c-greenCopy"><div align="center"><span id="iprogress"></span></div></td>
         </tr>
       </table></td>
     </tr>
   </table>
 </div>
 <form name="form1" method="post" action="/Reports/Tracking/JobTracking/JobTracking.asp">
   <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search Job Schedule </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Job Number</span>
    </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Job Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate" value="<% =formatdate(fromdate,application("shortdateformat"))%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate" value="<% =formatdate(todate,application("shortdateformat"))%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<form name="checkform" method="post" action="/Reports/Tracking/JobSchedule/ScheduleJobList1.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy">Browse Job Schedule of <%=schedule_name%></td>
</tr>
<tr>
  <td height="20" colspan="6" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobTracking_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span>
        </div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="6">&nbsp;</td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <%if admin=true then%>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JS.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=JS.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Quantity</div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Start Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td class="t-t-Borrow"><div align="center">Complete Time </div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
  <%if admin=true then%>
    <%end if%>
    <td height="20"><div align="center"><%= rs("JOB_NUMBER") %></div></td>
    <td><div align="center">
      <input name="jobnumber<%=i%>" type="hidden" id="jobnumber<%=i%>" value="<%=rs("JOB_NUMBER")%>">
      <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">
      <select name="part_number_id<%=i%>" id="part_number_id<%=i%>">
	  <option value=""></option>
	  <%=ubound(PART_ID)%>
	  <%for j=0 to ubound(PART_ID)%>
	  <option value="<%=PART_ID(j)%>" <%if rs("PART_NUMBER_ID")=PART_ID(j) then%>selected<%end if%>><%=PART_NAME(j)%></option>
	  <%next%>
      </select>
    </div></td>
    <td height="20"><div align="center"><%= rs("QUANTITY") %>&nbsp;</div></td>
    <td><div align="center"><% =formatdate(rs("START_TIME"),application("longdateformat"))%>
      &nbsp;</div></td>
    <td><div align="center"><% =formatdate(rs("COMPLETE_TIME"),application("longdateformat"))%>
      &nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
<tr>
  <td height="20" colspan="6"><span id="_PercentBar"></span>&nbsp;</td>
</tr>
<%if admin=true then%>
<tr>
  <td height="20" colspan="6"><div align="center">
      <input name="weekindex" type="hidden" id="weekindex" value="<%=weekindex%>">
      <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input name="Button1" type="button" id="Button1" value="Check All" onClick="checkall()">
&nbsp;
<input name="Button2" type="button" id="Button2" value="Uncheck All" onClick="uncheckall()">
&nbsp;
<input type="submit" name="Submit" value="Update" <%if admin=false then%>disabled<%end if%>>
&nbsp;
<input name="Reset" type="reset" id="Reset" value="Reset">
  </div></td>
  </tr>
<%
end if
else
%>
  <tr>
    <td height="20" colspan="6"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
</form>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->