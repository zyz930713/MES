<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetJobSetStartQuantity.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetJobStationGoodQuantity.asp" -->
<!--#include virtual="/Functions/GetJobStationDefectCodeNew.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<!--#include virtual="/Functions/GetJobFinalDefectQty.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
set rst=server.CreateObject("adodb.recordset")
set rst1=server.CreateObject("adodb.recordset")
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
fromdate=dateadd("d",-6,date())
end if
if todate="" then
todate=date()
end if

if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and lower(J.PART_NUMBER_TAG ) like '%"&lcase(partnumber)&"%'"
end if
if fromdate<>"" then
where=where&" and J.CLOSE_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.CLOSE_TIME<=to_date('"&todate&"','yyyy-mm-dd')+1"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Yield/JobYield/JobYield.asp"
FactoryRight "P."
SQL="select 1,J.JOB_NUMBER,J.SHEET_NUMBER,J.PART_NUMBER_TAG,J.START_TIME,J.CLOSE_TIME,J.FIRST_STATION_ID,J.LAST_STATION_ID,J.STATIONS_INDEX,J.JOB_START_QUANTITY,J.JOB_GOOD_QUANTITY,J.JOB_DEFECTCODE_QUANTITY,J.JOB_ASSEMBLY_YIELD,JOB_TYPE,P.PART_PER_FRAME as P_PER_FRAME from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.STATUS=1 "&where&order
'factorywhereoutsideand&order
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
</head>

<body onLoad="language_page()">
<form name="form1" method="post" action="/Reports/Yield/JobYield/JobYield.asp">
<table width="99%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
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
<table width="99%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="8" class="t-c-greenCopy">Browse Job Yield &nbsp;&nbsp;</td>
</tr>
<tr>
  <td height="20" colspan="8" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy">User: 
          <% =session("User") %></td>
		  <!--
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
		-->
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="8"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td width="3%" height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td width="8%" height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td width="8%" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Part Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td width="6%" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.CLOSE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Close Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.CLOSE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"> </div></td>
  <td width="10%" height="20" class="t-t-Borrow"><div align="center">Start Quantity</div></td>
  <td width="25%" class="t-t-Borrow"><div align="center">Defect Code </div></td>
  <td width="3%" class="t-t-Borrow"><div align="center">Total Defect Code</div></td>
  <td width="5%" class="t-t-Borrow"><div align="center"> Yield </div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")

dim thisjobnumber()
dim thisjobtype()
dim thissheetnumber()
dim thispartnumbertag()
dim thisstart_time()
dim thisclose_time()
dim thisstations_index()
dim thisjobstartquantity()
dim thisjobtotaldefectcodequantity()
dim thisjobgoodquantity()
dim thisjob_assembly_yield()
redim thisjobnumber(0)
redim thisjobtype(0)
redim thissheetnumber(0)
redim thispartnumbertag(0)
redim thisstart_time(0)
redim thisclose_time(0)
redim thisstations_index(0)
redim thisjobstartquantity(0)
redim thisjobtotaldefectcodequantity(0)
redim thisjobgoodquantity(0)
redim thisjob_assembly_yield(0)
total_jobstartquantity=0
total_jobtotaldefectcodequantity=0
total_jobgoodquantity=0
final_assembly_yield=0
current_jobnumber=rs("JOB_NUMBER")
t=1
while not rs.eof and i<=rs.pagesize 
	STATIONS_INDEX=rs("STATIONS_INDEX")	
	if rs("JOB_START_QUANTITY")<>"" then
		jobstartquantity=cint(rs("JOB_START_QUANTITY"))
	else
		jobstartquantity=0
	end if
	if rs("P_PER_FRAME")<>"0" then
		part_per_frame=rs("P_PER_FRAME")
	else 
		part_per_frame="1"
	end if
	jobstartnum=jobstartquantity*cint(part_per_frame)
	
	jobtotaldefectcodequantity=csng(rs("JOB_DEFECTCODE_QUANTITY"))
	jobgoodquantity=csng(rs("JOB_GOOD_QUANTITY"))
	'if cint(rs("JOB_ASSEMBLY_YIELD"))=0 then
'	jobstartquantity=getJobActionValue(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("FIRST_STATION_ID"),"AC00000065")
'	jobtotaldefectcodequantity=GetJobFinalDefectQty(rs("JOB_NUMBER"),rs("SHEET_NUMBER"))
'	jobgoodquantity=getJobStationGoodQuantity(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("LAST_STATION_ID"))
		if jobstartquantity<>0 then
		job_assembly_yield=cint(JobGoodQuantity)/jobstartquantity*cint(part_per_frame)
		else
		job_assembly_yield=0
		end if
	SQLU="update JOB set JOB_START_QUANTITY="&jobstartquantity&",JOB_GOOD_QUANTITY="&jobgoodquantity&",JOB_DEFECTCODE_QUANTITY="&jobtotaldefectcodequantity&",JOB_ASSEMBLY_YIELD="&job_assembly_yield&" where JOB_NUMBER='"&rs("JOB_NUMBER")&"' and SHEET_NUMBER='"&rs("SHEET_NUMBER")&"'"
	rsU.open SQLU,conn,1,3
	'else
	'jobstartquantity=rs("JOB_START_QUANTITY")
	'jobtotaldefectcodequantity=rs("JOB_DEFECTCODE_QUANTITY")
	'jobgoodquantity=rs("JOB_GOOD_QUANTITY")
	'job_assembly_yield=csng(rs("JOB_ASSEMBLY_YIELD"))
	'end if
	
	if rs("JOB_NUMBER")=current_jobnumber then 'Get total parameters' value of all the same jobs
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thisjobtype(UBound(thisjobtype))=rs("JOB_TYPE")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisclose_time(UBound(thisclose_time))=rs("CLOSE_TIME")	
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thisjobtype(UBound(thisjobtype)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisclose_time(UBound(thisclose_time)+1)
	ReDim Preserve thisstations_index(UBound(thisstations_index)+1)
	ReDim Preserve thisjobstartquantity(UBound(thisjobstartquantity)+1)
	ReDim Preserve thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity)+1)
	ReDim Preserve thisjobgoodquantity(UBound(thisjobgoodquantity)+1)
	ReDim Preserve thisjob_assembly_yield(UBound(thisjob_assembly_yield)+1)
	total_jobstartquantity=total_jobstartquantity+cint(jobstartquantity)
	total_jobtotaldefectcodequantity=total_jobtotaldefectcodequantity+clng(jobtotaldefectcodequantity)
	total_jobgoodquantity=total_jobgoodquantity+cint(jobgoodquantity)
	else
	 if total_jobstartquantity<>0 then
	 	PER_jobyield=formatpercent(clng(total_jobgoodquantity)/(clng(total_jobstartquantity)*cint(part_per_frame)),2)
	 else
	 	PER_jobyield=0
	 end if
	%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')"><input name="tabstatus<%=current_jobnumber%>" type="hidden" value="1"><img src="/Images/Treeimg/minus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td height="20"><%thisJobSetStartQuantity=getJobSetStartQuantity(current_jobnumber)%>
      <div align="center" <%if clng(thisJobSetStartQuantity)<>clng(total_jobstartquantity) then%>class="strongred"<%end if%>><%=thisJobSetStartQuantity%><%if part_per_frame<>"1" then%>*<%=part_per_frame%><%end if%>(<%=total_jobstartquantity%><%if part_per_frame<>"1" then%>*<%=part_per_frame%><%end if%>)</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=total_jobtotaldefectcodequantity%>&nbsp;</div></td>
    <td><div align="center">
        
	<%=PER_jobyield%>
	&nbsp;</div></td>
  </tr>
	<tbody id="tab<%=current_jobnumber%>">
<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
			if cint(thisjobstartquantity(u))<>0 then
			 PER_jobSTQTY=formatpercent(cint(thisjobgoodquantity(u))/(cint(thisjobstartquantity(u))*cint(part_per_frame)),2)
			 else
			 	PER_jobSTQTY=0
			 end if
%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>&jobtype=<%=thisjobtype(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td height="20"><div align="center"><%= thispartnumbertag(u) %>&nbsp;</div></td>
		<td><div align="center"><% =formatdate(thisclose_time(u),application("longdateformat"))%></div></td>
		<td height="20"><div align="center"><%=thisjobstartquantity(u)*cint(part_per_frame)%></div></td>
		<td><div align="center">
		<%=getJobStationDefectCode(thisjobnumber(u),thissheetnumber(u),thisstations_index(u))%>
		</div></td>
	  <td><div align="center"><%=thisjobtotaldefectcodequantity(u)%>&nbsp;</div></td>
		<td><div align="center">
		<%=PER_jobSTQTY%>&nbsp;</div></td>
		
	  </tr>
	  
<%		next%>
	</tbody>
<%
	redim thisjobnumber(0)
	redim thisjobtype(0)
	redim thissheetnumber(0)
	redim thispartnumbertag(0)
	redim thisstart_time(0)
	redim thisclose_time(0)	
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
	thisjobtype(UBound(thisjobtype))=rs("JOB_TYPE")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisclose_time(UBound(thisstart_time))=rs("CLOSE_TIME")	
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thisjobtype(UBound(thisjobtype)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisclose_time(UBound(thisclose_time)+1)	
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
current_jobnumber=rs("JOB_NUMBER")
rs.movenext
wend
%>
	<tr class="t-b-blue">
  <td height="20"><div align="center"><% =t%></div></td>
    <td height="20"><div align="center"><span style="cursor:hand" onClick="tabplus('<%=current_jobnumber%>')"><input name="tabstatus<%=current_jobnumber%>" type="hidden" value="1"><img src="/Images/Treeimg/minus.gif" name="tabimg<%=current_jobnumber%>" width="9" height="9">&nbsp;<%= current_jobnumber %></span></div></td>
    <td height="20"><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td height="20"><%thisJobSetStartQuantity=getJobSetStartQuantity(current_jobnumber)%>
      <div align="center" <%if clng(thisJobSetStartQuantity)<>clng(total_jobstartquantity) then%>class="strongred"<%end if%>><%=thisJobSetStartQuantity%><%if part_per_frame<>"1" then%>*<%=part_per_frame%><%end if%>(<%=total_jobstartquantity%><%if part_per_frame<>"1" then%>*<%=part_per_frame%><%end if%>)</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center"><%=total_jobtotaldefectcodequantity%>&nbsp;</div></td>
    <td><div align="center">
	<%if total_jobstartquantity<>0 then
	final_assembly_yield=cint(total_jobgoodquantity)/(cint(total_jobstartquantity)*cint(part_per_frame))
	else
	final_assembly_yield=0
	end if%>
	<%=formatpercent(final_assembly_yield,2)%>&nbsp;</div></td>
  </tr>
  <tbody id="tab<%=current_jobnumber%>">
	<%
		for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
%>
	<tr>
	  <td height="20"><div align="center"><%=u+1%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=thisjobnumber(u)%>&sheetnumber=<%=thissheetnumber(u)%>&jobtype=<%=thisjobtype(u)%>" target="_blank"><%= thisjobnumber(u) %>-<%=repeatstring(thissheetnumber(u),"0",3)%></a></div></td>
		<td height="20"><div align="center"><%= thispartnumbertag(u) %>&nbsp;</div></td>
		<td><div align="center"><% =formatdate(thisclose_time(u),application("longdateformat"))%></div></td>
		<td height="20"><div align="center"><%=thisjobstartquantity(u)*cint(part_per_frame)%></div></td>
		<td><div align="center">
		<%=getJobStationDefectCode(thisjobnumber(u),thissheetnumber(u),thisstations_index(u))%>
		</div></td>
		<td><div align="center"><%=thisjobtotaldefectcodequantity(u)%>&nbsp;</div></td>
		<td><div align="center"><%=formatpercent(cint(thisjobgoodquantity(u))/(cint(thisjobstartquantity(u))*cint(part_per_frame)),2)%>&nbsp;</div></td>
	  </tr>
<%
		next%>
  </tbody>
<%
else
%>
  <tr>
    <td height="20" colspan="8"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsU=nothing
set rst=nothing
set rst1=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->