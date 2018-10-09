<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetJobStationGoodQuantity.asp" -->
<!--#include virtual="/Functions/GetJobStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER asc"
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
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Yield/JobYield/JobYield.asp"
SQL="select 1,J.JOB_NUMBER,J.STATIONS_INDEX,J.STATIONS_TRANSACTION from JOB J where J.STATUS=1 "&where&order
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
<form name="form1" method="post" action="/Reports/Yield/JobYield/JobYield.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search Job </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Job Number</span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Job Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate2" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate2" value="<%=todate%>" size="10">
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
  <td height="20" colspan="3" class="t-c-greenCopy">Browse Job Yield </td>
</tr>
<tr>
  <td height="20" colspan="3" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="3"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Stations</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
	astation=split(rs("STATIONS_INDEX"),",")
	atransaction()=splt(rs(""),",")
	if cint(rs("JOB_ASSEMBLY_YIELD"))=0 then
	jobstartquantity=getJobActionValue(rs("JOB_NUMBER"),rs("FIRST_STATION_ID"),"AC00000097")
	jobtotaldefectcodequantity=getJobTotalDefectCodeQuantity(rs("JOB_NUMBER"))
	jobgoodquantity=getJobStationGoodQuantity(rs("JOB_NUMBER"),rs("LAST_STATION_ID"))
	job_assembly_yield=cint(JobGoodQuantity)/cint(jobstartquantity)
	SQLU="update JOB set JOB_START_QUANTITY="&jobstartquantity&",JOB_GOOD_QUANTITY="&jobgoodquantity&",JOB_DEFECTCODE_QUANTITY="&jobtotaldefectcodequantity&",JOB_ASSEMBLY_YIELD="&job_assembly_yield&" where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
	rsU.open SQLU,conn,1,3
	else
	jobstartquantity=rs("JOB_START_QUANTITY")
	jobtotaldefectcodequantity=rs("JOB_DEFECTCODE_QUANTITY")
	jobgoodquantity=rs("JOB_GOOD_QUANTITY")
	job_assembly_yield=csng(rs("JOB_ASSEMBLY_YIELD"))
	end if
%>
<tr>
  <td><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%>
  </div></td>
  <td><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("JOB_NUMBER") %></a></div></td>
  <td height="20"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <%for j=0 to ubound(astation)
  		if atraction(j)="0" then
		defectcode_quantity=getDefectCode(null)%>
    <tr>
      <td colspan="3"><div align="center"><%=stationname%></div></td>
      </tr>
    <tr>
      <td><div align="center">Initial</div></td>
      <td><div align="center">Defectcode</div></td>
      <td><div align="center">Good</div></td>
    </tr>
    <tr>
      <td><div align="center"><%=defectcode_quantity%>&nbsp;</div></td>
      <td><div align="center"><%=defectcode_quantity%></div></td>
      <td><div align="center"><%=defectcode_quantity%></div></td>
    </tr>
  </table>    </td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="3"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->