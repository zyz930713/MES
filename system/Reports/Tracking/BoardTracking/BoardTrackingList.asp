<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobTracking/JobTrackingCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
yearindex=trim(request("yearindex"))
weekindex=trim(request("weekindex"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by NID desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if yearindex<>"" then
where=where&" and YEAR_INDEX='"&yearindex&"'"
end if
if weekindex<>"" then
where=where&" and WEEK_INDEX='"&weekindex&"'"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=todate('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&yearindex="&yearindex&"&weekindex="&weekindex&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Tracking/BoardTracking/BoardTrackingList.asp"
SQL="select * from JOB_BOARD_LIST where NID is not null "&where&order
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
</head>

<body>
<form name="form1" method="post" action="/Reports/Tracking/BoardTracking/BoardTrackingList.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search Board Tracking Reports </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Year</span> </td>
    <td height="20"><input name="yearindex" type="text" id="yearindex" value="<%=yearindex%>">

    </td>
    <td>Week</td>
    <td><input name="weekindex" type="text" id="weekindex" value="<%=weekindex%>"></td>
    <td>Report Time</td>
    <td>From
        <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
      <input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
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
	  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy">Browse Board Tracking </td>
  </tr>
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%">User:
              <% =session("User") %></td>
          <td width="50%"><div align="right">
              <%if admin=true then%>
              <input name="Generate" type="button" id="Generate" value="Generate Board Tracking Report At Once" onClick="javascript:location.href='GenerateBoardTracking.asp?path=<%=path%>&query=<%=query%>'">
              <%end if%>
          </div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
    <%if admin=true then%>
    <td class="t-t-Borrow"><div align="center">Delete</div></td>
    <%end if%>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WEEK_INDEX&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Week of Schedule<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WEEK_INDEX&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=YEAR_INDEX&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Year of Schedule<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=YEAR_INDEX&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=TRACKING_NAME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=TRACKING_NAME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=CREATOR_CODE&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Creator<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=CREATOR_CODE&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=UPDATE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Last Update Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=UPDATE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center">Excel File </div></td>
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
    <%if admin=true then%>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:location.href='DeleteBoardTracking.asp?id=<%=rs("NID")%>&board_name=<%=rs("TRACKING_NAME")%>&excelfilename=<%=rs("EXCEL_FILENAME")%>&path=<%=path%>&query=<%=query%>'">Delete</span></div></td>
    <%end if%>
    <td height="20"><div align="center"><%= rs("WEEK_INDEX") %></div></td>
    <td><div align="center"><%= rs("YEAR_INDEX") %></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('BoardTracking.asp?board_id=<%=rs("NID")%>&board_name=<%=rs("TRACKING_NAME")%>&board_report_time=<%=rs("REPORT_TIME")%>&fromdate=<%=rs("WEEK_START_DAY")%>&todate=<%=rs("WEEK_END_DAY")%>')"><%= rs("TRACKING_NAME") %></span>&nbsp;</div></td>
    <td><div align="center"><%= rs("CREATOR_CODE") %></div></td>
    <td><div align="center">
        <% =formatdate(rs("REPORT_TIME"),application("longdateformat"))%>
    </div></td>
    <td><div align="center">
        <% =formatdate(rs("UPDATE_TIME"),application("longdateformat"))%>
    </div></td>
    <td><div align="center"><a href="/Reports/Tracking/BoardTracking/EXCELs/<%=rs("EXCEL_FILENAME")%>" target="_blank"><img src="/Images/EXCELs.gif" width="15" height="15" border="0"></a></div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->