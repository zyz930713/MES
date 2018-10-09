<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetJobSetStartQuantity.asp" -->
<!--#include virtual="/Functions/GetPartTargetYield.asp" -->
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
order=" order by J.PART_NUMBER_TAG asc,J.JOB_NUMBER asc,J.SHEET_NUMBER asc"
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
where=where&" and J.CLOSE_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Yield/PartYield/PartYield.asp"
FactoryRight "P."
SQL="select 1,J.LINE_NAME,J.JOB_NUMBER,J.SHEET_NUMBER,J.PART_NUMBER_ID,J.START_TIME,J.FIRST_STATION_ID,J.LAST_STATION_ID,J.PART_NUMBER_TAG,J.STATIONS_INDEX,J.JOB_START_QUANTITY,J.JOB_GOOD_QUANTITY,J.JOB_DEFECTCODE_QUANTITY,J.JOB_ASSEMBLY_YIELD from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.STATUS=1 "&where&factorywhereoutsideand&order
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
<script language=JavaScript src="/Functions/TablePlus.js" type=text/javascript></script>
</head>

<body>
<form name="form1" method="post" action="/Reports/Yield/PartYield/PartYield.asp">
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
  <td height="20" colspan="9" class="t-c-greenCopy">Browse Part Yield </td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy">User: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('PartYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="9">&nbsp;</td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Part</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Start Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Start Quantity</div></td>
  <td class="t-t-Borrow"><div align="center">Station Defect Code </div></td>
  <td class="t-t-Borrow"><div align="center">Total Defect Code</div></td>
  <td class="t-t-Borrow"><div align="center">Target Yield</div></td>
  <td class="t-t-Borrow"><div align="center"> Yield </div></td>
  </tr>
<%
i=1
if not rs.eof then

dim thisjobnumber()
dim thissheetnumber()
dim thispartnumbertag()
dim thisstart_time()
dim thisstations_index()
dim thisjobstartquantity()
dim thisjobtotaldefectcodequantity()
dim thisjobgoodquantity()
dim thisjob_assembly_yield()
redim thisjobnumber(0)
redim thissheetnumber(0)
redim thispartnumbertag(0)
redim thisstart_time(0)
redim thisstations_index(0)
redim thisjobstartquantity(0)
redim thisjobtotaldefectcodequantity(0)
redim thisjobgoodquantity(0)
redim thisjob_assembly_yield(0)
total_jobstartquantity=0
total_jobtotaldefectcodequantity=0
total_jobgoodquantity=0
final_assembly_yield=0
current_part_tag=rs("PART_NUMBER_TAG")
current_part_id=rs("PART_NUMBER_ID")
t=1
while not rs.eof
	if cint(rs("JOB_ASSEMBLY_YIELD"))=0 then
	jobstartquantity=getJobActionValue(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("FIRST_STATION_ID"),"AC00000097")
	jobtotaldefectcodequantity=getJobTotalDefectCodeQuantity(rs("JOB_NUMBER"),rs("SHEET_NUMBER"))
	jobgoodquantity=getJobStationGoodQuantity(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("LAST_STATION_ID"))
		if jobstartquantity<>0 then
		job_assembly_yield=cint(JobGoodQuantity)/cint(jobstartquantity)
		else
		job_assembly_yield=0
		end if
	SQLU="update JOB set JOB_START_QUANTITY="&jobstartquantity&",JOB_GOOD_QUANTITY="&jobgoodquantity&",JOB_DEFECTCODE_QUANTITY="&jobtotaldefectcodequantity&",JOB_ASSEMBLY_YIELD="&job_assembly_yield&" where JOB_NUMBER='"&rs("JOB_NUMBER")&"' and SHEET_NUMBER='"&rs("SHEET_NUMBER")&"'"
	rsU.open SQLU,conn,1,3
	else
	jobstartquantity=rs("JOB_START_QUANTITY")
	jobtotaldefectcodequantity=rs("JOB_DEFECTCODE_QUANTITY")
	jobgoodquantity=rs("JOB_GOOD_QUANTITY")
	job_assembly_yield=csng(rs("JOB_ASSEMBLY_YIELD"))
	end if
	
	if rs("PART_NUMBER_TAG")=current_part_tag then 'Get total parameters' value of all the same jobs
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisstations_index(UBound(thisstations_index)+1)
	ReDim Preserve thisjobstartquantity(UBound(thisjobstartquantity)+1)
	ReDim Preserve thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity)+1)
	ReDim Preserve thisjobgoodquantity(UBound(thisjobgoodquantity)+1)
	ReDim Preserve thisjob_assembly_yield(UBound(thisjob_assembly_yield)+1)
	total_jobstartquantity=total_jobstartquantity+cint(jobstartquantity)
	total_jobtotaldefectcodequantity=total_jobtotaldefectcodequantity+cint(jobtotaldefectcodequantity)
	total_jobgoodquantity=total_jobgoodquantity+cint(jobgoodquantity)
	else
	%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td><span style="cursor:hand" onClick="tabplus('<%=replace(current_part_tag,"-","_")%>')">
      <input name="tabstatus<%=replace(current_part_tag,"-","_")%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=replace(current_part_tag,"-","_")%>" width="9" height="9">&nbsp;<%= current_part_tag %></span></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td height="20">
      <div align="center"><%=total_jobstartquantity%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=total_jobtotaldefectcodequantity%>&nbsp;</div></td>
    	<%if total_jobstartquantity<>0 then
	final_assembly_yield=cint(total_jobgoodquantity)/cint(total_jobstartquantity)
	else
	final_assembly_yield=0
	end if
	this_target_yield=getPartTargetYield(current_part_id)%>
	<td><div align="center"><span <%if csng(final_assembly_yield)<csng(this_target_yield)/100 then%>class="red"<%end if%>><%=this_target_yield%>%</span></div></td>
    <td><div align="center">
	<%=formatpercent(final_assembly_yield,2)%>&nbsp;</div></td>
  </tr>
	<tbody id="tab<%=replace(current_part_tag,"-","_")%>" style="display:none">
<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td>&nbsp;</td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td><div align="center"><% =formatdate(thisstart_time(u),application("longdateformat"))%></div></td>
		<td height="20"><div align="center"><%=thisjobstartquantity(u)%></div></td>
		<td><div align="center"><%=getJobStationDefectCode(thisjobnumber(u),thissheetnumber(u),thisstations_index(u))%></div></td>
		<td><div align="center"><%=thisjobtotaldefectcodequantity(u)%>&nbsp;</div></td>
		<td><div align="center">&nbsp;</div></td>
		<td><div align="center"><%=formatpercent(thisjob_assembly_yield(u),2)%>&nbsp;</div></td>
	  </tr>
	  
<%		next%>
	</tbody>
<%
	redim thisjobnumber(0)
	redim thissheetnumber(0)
	redim thispartnumbertag(0)
	redim thisstart_time(0)
	redim thisstations_index(0)
	redim thisjobstartquantity(0)
	redim thisjobtotaldefectcodequantity(0)
	redim thisjobgoodquantity(0)
	redim thisjob_assembly_yield(0)
	total_jobstartquantity=0
	total_jobtotaldefectcodequantity=0
	total_jobgoodquantity=0
	final_assembly_yield=0
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisstations_index(UBound(thisstations_index)+1)
	ReDim Preserve thisjobstartquantity(UBound(thisjobstartquantity)+1)
	ReDim Preserve thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity)+1)
	ReDim Preserve thisjobgoodquantity(UBound(thisjobgoodquantity)+1)
	ReDim Preserve thisjob_assembly_yield(UBound(thisjob_assembly_yield)+1)
	total_jobstartquantity=total_jobstartquantity+cint(jobstartquantity)
	total_jobtotaldefectcodequantity=total_jobtotaldefectcodequantity+cint(jobtotaldefectcodequantity)
	total_jobgoodquantity=total_jobgoodquantity+cint(jobgoodquantity)
	t=t+1
	end if
i=i+1
current_part_tag=rs("PART_NUMBER_TAG")
current_part_id=rs("PART_NUMBER_ID")
rs.movenext
wend
%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td><span style="cursor:hand" onClick="tabplus('<%=replace(current_part_tag,"-","_")%>')">
      <input name="tabstatus<%=replace(current_part_tag,"-","_")%>" type="hidden" value="0">
      <img src="/Images/Treeimg/plus.gif" name="tabimg<%=replace(current_part_tag,"-","_")%>" width="9" height="9">&nbsp;<%= current_part_tag %></span></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td height="20">
      <div align="center"><%=total_jobstartquantity%></div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=total_jobtotaldefectcodequantity%>&nbsp;</div></td>
    	<%if total_jobstartquantity<>0 then
	final_assembly_yield=cint(total_jobgoodquantity)/cint(total_jobstartquantity)
	else
	final_assembly_yield=0
	end if
	this_target_yield=getPartTargetYield(current_part_id)%>
	<td><div align="center"><span <%if csng(final_assembly_yield)<csng(this_target_yield)/100 then%>class="red"<%end if%>><%=this_target_yield%>%</span></div></td>
    <td><div align="center">
	<%=formatpercent(final_assembly_yield,2)%>&nbsp;</div></td>
  </tr>
  <tbody id="tab<%=replace(current_part_tag,"-","_")%>" style="display:none">
	<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td>&nbsp;</td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td><div align="center"><% =formatdate(thisstart_time(u),application("longdateformat"))%></div></td>
		<td height="20"><div align="center"><%=thisjobstartquantity(u)%></div></td>
		<td><div align="center"><%=getJobStationDefectCode(thisjobnumber(u),thissheetnumber(u),thisstations_index(u))%></div></td>
		<td><div align="center"><%=thisjobtotaldefectcodequantity(u)%>&nbsp;</div></td>
		<td><div align="center">&nbsp;</div></td>
		<td><div align="center"><%=formatpercent(thisjob_assembly_yield(u),2)%>&nbsp;</div></td>
    </tr>
<%
		next%>
  </tbody>
<%
else
%>
  <tr>
    <td height="20" colspan="9"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->