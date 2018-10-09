<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<%
jobnumber=trim(request("jobnumber"))
sheetnumber=trim(request("sheetnumber"))
jobtype=trim(request("jobtype"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
factory=trim(request("factory"))
seriesgroup=trim(request("seriesgroup"))
timespan=trim(request("timespan"))
jobstatus=trim(request("jobstatus"))
currentstation=trim(request("currentstation"))
fromdate=request("fromdate")
todate=request("todate")
close_fromdate=request("close_fromdate")
close_todate=request("close_todate")
ordername=request("ordername")
ordertype=request("ordertype")
DisplayStyle=request("DisplayStyle")

fromtime1=request("fromtime1")
if isnull(fromtime1) or fromtime1=""  then
	fromtime1="00:00:00"
end if
fromtime2=request("fromtime2")
if isnull(fromtime2) or fromtime2=""  then
	fromtime2="23:59:59"
end if

endtime1=request("endtime1")
if isnull(endtime1) or endtime1="" then
	endtime1="00:00:00"
end if

endtime2=request("endtime2")
if isnull(endtime2) or endtime2="" then
	endtime2="23:59:59"
end if


if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc,J.SHEET_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER = '"&jobnumber&"'"
end if
if sheetnumber<>"" then
where=where&" and J.SHEET_NUMBER='"&sheetnumber&"'"
end if
if jobtype<>"" then
where=where&" and J.JOB_TYPE='"&jobtype&"'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG = '"&partnumber&"'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME) like '%"&lcase(line)&"%'"
end if
if factory<>"" then
where=where&" and P.FACTORY_ID='"&factory&"'"
end if
if seriesgroup<>"" then
SQL="Select INCLUDED_SYSTEM_ITEMS from SERIES_GROUP where NID='"&seriesgroup&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	seriesgrouplist=rs("INCLUDED_SYSTEM_ITEMS")
end if
rs.close
where=where&" and J.PART_NUMBER_TAG in (select * from table(CHAR_SPLIT('"&seriesgrouplist&"',',')))"
end if
if jobstatus<>"" then
where=where&" and J.STATUS="&jobstatus
end if
if currentstation<>"" then
where=where&" and S.MOTHER_STATION_ID='"&currentstation&"'"
end if
if timespan<>"" then
where=where&" and J.CYCLE_TIME>="&timespan*60
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate& " " &fromtime1 &"','yyyy-mm-dd hh24:mi:ss')"
elseif close_fromdate="" and close_todate="" and todate="" then
fromdate=Format_Time(dateadd("d",-1,date()),2)
'where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
where=where&" and J.START_TIME>=to_date('"&fromdate& " " &fromtime1 &"','yyyy-mm-dd hh24:mi:ss')"
end if
if todate<>"" then
'where=where&" and J.START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
where=where&" and J.START_TIME<=to_date('"&todate& " " &fromtime2 &"','yyyy-mm-dd hh24:mi:ss')"
end if
if close_fromdate<>"" then
'where=where&" and J.CLOSE_TIME>=to_date('"&close_fromdate&"','yyyy-mm-dd')"
where=where&" and J.CLOSE_TIME>=to_date('"&close_fromdate& " " &endtime1 &"','yyyy-mm-dd hh24:mi:ss')"
end if
if close_todate<>"" then
'where=where&" and J.CLOSE_TIME<=to_date('"&close_todate&"','yyyy-mm-dd')"
where=where&" and J.CLOSE_TIME<=to_date('"&close_todate& " " &endtime2 &"','yyyy-mm-dd hh24:mi:ss')"
end if
pagepara="&jobnumber="&jobnumber&"&sheetnumber="&sheetnumber&"&jobtype="&jobtype&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&seriesgroup="&seriesgroup&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&timespan="&timespan&"&fromdate="&fromdate&"&todate="&todate&"&close_fromdate="&close_fromdate&"&close_todate="&close_todate&"&ordername="&ordername&"&ordertype="&ordertype &"&DisplayStyle="&DisplayStyle
FactoryRight "P."
SQL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID LEFT JOIN STATION S ON J.CURRENT_STATION_ID=S.NID where 1=1 "&where&factorywhereoutsideand&order
'response.Write(sql)
dim StationIndexSQL
StationIndexSQL="select DISTINCT J.STATIONS_INDEX  from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID  LEFT JOIN STATION S ON J.CURRENT_STATION_ID=S.NID where J.JOB_NUMBER is not null"&where&factorywhereoutsideand
session("StationIndexSQL")=StationIndexSQL

dim SQLEXCEL
SQLEXCEL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION,"
SQLEXCEL=SQLEXCEL & " (select station_start_quantity  From job_stations where job_number=J.JOB_NUMBER and sheet_number=J.SHEET_NUMBER and station_id=j.previous_station_id) as Current_Qty"
SQLEXCEL=SQLEXCEL & " from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null"&where&factorywhereoutsideand&order
 

'response.End()
rs.open SQL,conn,1,3
 
'response.write sql
session("SQL")=SQLEXCEL
session("SQLStation")="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null"&where&factorywhereoutsideand&" order by J.PART_NUMBER_ID,J.JOB_NUMBER,J.SHEET_NUMBER"

'日期格式化
Function Format_Time(s_Time, n_Flag) 
Dim y, m, d, h, mi, s 
Format_Time = "" 
If IsDate(s_Time) = False Then Exit Function 
y = cstr(year(s_Time)) 
m = cstr(month(s_Time)) 
If len(m) = 1 Then m = "0" & m 
d = cstr(day(s_Time)) 
If len(d) = 1 Then d = "0" & d 
h = cstr(hour(s_Time)) 
If len(h) = 1 Then h = "0" & h 
mi = cstr(minute(s_Time)) 
If len(mi) = 1 Then mi = "0" & mi 
s = cstr(second(s_Time)) 
If len(s) = 1 Then s = "0" & s 
Select Case n_Flag 
Case 1 
' yyyy-mm-dd hh:mm:ss 
Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
Case 2 
' yyyy-mm-dd 
Format_Time = y & "-" & m & "-" & d 
Case 3 
' hh:mm:ss 
Format_Time = h & ":" & mi & ":" & s 
Case 4 
' yyyy年mm月dd日 
Format_Time = y & "年" & m & "月" & d & "日" 
Case 5 
' yyyymmdd 
Format_Time = y & m & d 
case 6 
'yyyymmddhhmmss 
format_time= y & m & d & h & mi & s 
End Select 
End Function 


'response.Write(session("SQLStation") & "<br>")
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Job/SubJobs/Lan_Job.asp" -->
</head>
<body onLoad="language();language_page();language_jobnote()">
<form action="/Job/SubJobs/Job.asp" method="get" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_Search"></span></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchJobNumber"></span></td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">
        -
        <input name="sheetnumber" type="text" id="sheetnumber" value="<%=sheetnumber%>" size="3" maxlength="3">
	  </td>
	  <td height="20"><span id="inner_SearchLineName"></span></td>
      <td height="20"><input name="line" type="text" id="line" value="<%=line%>"></td>
	  <td><span id="inner_SearchJobType"></span></td>
      <td><select name="jobtype" id="jobtype">
          <option value="" <%if jobtype="" then%>selected<%end if%>>All</option>
          <option value="N" <%if jobtype="N" then%>selected<%end if%>>Normal</option>
          <option value="R" <%if jobtype="R" then%>selected<%end if%>>Retest</option>
        </select>
      </td>	  
	  <td><span id="inner_SearchCurrentStation"></span></td>
      <td><select name="currentstation" id="currentstation">
          <option value="">All</option>
          <%=getStation_New(true,"OPTION",currentstation,""," order by S.STATION_NAME","","")%>
        </select></td>
 	  <td height="20"><span id="inner_SearchStatus"></span></td>
      <td height="20"><select name="jobstatus" id="jobstatus">
          <option value="">All</option>
          <option value="0" <%if jobstatus="0" then%>selected<%end if%>>Opened</option>
          <option value="2" <%if jobstatus="2" then%>selected<%end if%>>Paused</option>
          <option value="1" <%if jobstatus="1" then%>selected<%end if%>>Closed</option>
		  <!--
          <option value="3" <%if jobstatus="3" then%>selected<%end if%>>Locked/被锁定</option>
		  -->
          <option value="5" <%if jobstatus="5" then%>selected<%end if%>>Locked</option>
        </select></td>			        
    </tr>
    <tr>
	  <td><span id="inner_SearchPartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
	<!--
	  <td><span id="inner_SearchFactory"></span></td>
      <td><select name="factory" id="factory">
          <option value="">Factory</option>
          <%FactoryRight ""%>
          <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
        </select></td>
      <td><span id="inner_SearchCycleTime"></span></td>
      <td>&gt;=
        <input name="timespan" type="text" id="timespan" value="<%=timespan%>" size="4">
        <span id="inner_SearchM"></span>h</td>
		-->
      <td height="20"><span id="inner_SearchJobStartTime"></span></td>
      <td height="20" colspan="3" ><span id="inner_SearchFrom"></span>
        <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
<!--						
        <select name="fromtime1" id="fromtime1">
          <option value="00:00:00" <% if fromtime1="00:00:00" then response.write "Selected" end if%>>00:00:00</option>
          <option value="14:30:00" <% if fromtime1="14:30:00" then response.write "Selected" end if%>>14:30:00</option>
          <option value="23:59:59" <% if fromtime1="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
        </select>
-->		
        &nbsp;<span id="inner_SearchTo"></span>
        <input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
        &nbsp;
<!--		
        <select name="fromtime2" id="fromtime2">
          <option value="00:00:00" <% if fromtime2="00:00:00" then response.write "Selected" end if%>>00:00:00</option>
          <option value="14:30:00" <% if fromtime2="14:30:00" then response.write "Selected" end if%>>14:30:00</option>
          <option value="23:59:59" <% if fromtime2="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
        </select>
-->			
      </td>
      <td><span id="inner_SearchJobCloseTime"></span></td>
      <td ><span id="inner_SearchFrom1"></span>
	  <input name="close_fromdate" type="text" id="close_fromdate" value="<%=close_fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.close_fromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
<!--						
        <select name="endtime1" id="endtime1">
          <option value="00:00:00" <% if endtime1="00:00:00" then response.write "Selected" end if%>>00:00:00</option>
          <option value="14:30:00" <% if endtime1="14:30:00" then response.write "Selected" end if%>>14:30:00</option>
          <option value="23:59:59" <% if endtime1="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
        </select>
-->		
        &nbsp;<span id="inner_SearchTo1"></span>
        <input name="close_todate" type="text" id="close_todate" value="<%=close_todate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar4Callback(date, month, year)
	{
	document.all.close_todate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
    </script>
<!--	
        <select name="endtime2" id="endtime2">
          <option value="00:00:00" <% if endtime2="00:00:00" then response.write "Selected" end if%>>00:00:00</option>
          <option value="14:30:00" <% if endtime2="14:30:00" then response.write "Selected" end if%>>14:30:00</option>
          <option value="23:59:59" <% if endtime2="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
        </select>
-->
      </td>
	  <td colspan="2">
	  <img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"> 
	  				  
	  </td>
    </tr>
  </table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_User"></span>:<% =session("User") %></td>
          <td ><div align="right">
			<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Job_Export.asp')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>
		  	&nbsp;
			<span style="cursor:hand" title="Export Detailed Data to EXCEL File" onClick="javascript:window.open('JobStation_Export.asp')"><img src="/Images/IconReportTable.gif" width="22" height="22"></span>
              </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
    <%if admin=true then%>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
    <%if DBA=true then%>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Delete"></span></div></td>
    <%end if
	  end if%>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_PartType"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_StartTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_CloseTime"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
	<!--
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CurrentStations"></span></div></td>
	-->
    <td class="t-t-Borrow"><div align="center"><span id="inner_Operators"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_LineLost"></span></div></td>
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize
	session("aerror")=rs("JOB_NUMBER")&"-"&rs("SHEET_NUMBER")
	station_text=""
	stations_index=rs("STATIONS_INDEX")
	part_stations_index=rs("P_STATIONS_INDEX")
	repeated_stations_sequence=rs("REPEATED_STATIONS_SEQUENCE")
	finished_stations_id=rs("FINISHED_STATIONS_ID")
	if not isnull(rs("OPENED_STATIONS_ID")) and rs("OPENED_STATIONS_ID")<>"" then
		opened_stations_id=left(rs("OPENED_STATIONS_ID"),len(rs("OPENED_STATIONS_ID"))-1)
	else
		opened_stations_id=""
	end if
	current_station_id=rs("CURRENT_STATION_ID")
	transaction_type=rs("STATIONS_TRANSACTION")
	if isnull(rs("STATIONS_INDEX")) or rs("STATIONS_INDEX")="" then
		stations_index=part_stations_index
		transaction_type=rs("P_STATIONS_TRANSACTION")
	end if
	if rs("STATIONS_ROUTINE")="1" then
		stations_index=opened_stations_id
	end if
%>
  <tr>
    <td height="20"><div align="center">
        <% =(csng(session("strpagenum"))-1)*recordsize+i%>
      </div></td>
    <td><div align="center">
        <%SubJobStatus rs("STATUS"),simg,aimg,text,ctext,alt,calt,apage%>
        <img src="/Images/<%=simg%>.gif"></div>
      <%
		if rs("STATUS")="0" then
			set rsStatus=server.CreateObject("adodb.recordset") 
			mCurrent_Status_ID=rs("current_station_id")
			SQL="SELECT * FROM job_stations where JOB_NUMBER='"+rs("JOB_NUMBER")+"' AND SHEET_NUMBER='"+cstr(rs("SHEET_NUMBER"))+"' and station_id='"+mCurrent_Status_ID+"'"
			rsStatus.open SQL,conn,1,3
			if(rsStatus.recordcount=0) then
				Response.write "WAIT"
			end if
			if(rsStatus.recordcount=1) then
				Response.write "RUN"
			end if
		end if 
		if rs("STATUS")="2" then
				Response.write "HOLD"
		end if 
	%>
    </td>
    <%if admin=true then%>
    <td height="20" class="red"><div align="center">
        <%
	  actionimg=split(aimg,",")
	  actionalt=split(alt,",")
	  actioncalt=split(calt,",")
	  actionpage=split(apage,",")
		  for j=0 to ubound(actionimg)%>
        <img style="cursor:hand" src="/Images/<%=actionimg(j)%>.gif" alt="click to <%=actionalt(j)%> job.点击<%=actioncalt(j)%>工单" onClick="javascript:if('<%=actionalt(j)%>'=='pause' || '<%=actionalt(j)%>'=='Start'){alert('Please use the corresponding function to do these actions:Pause Job==>Hold Job; Start Job==>Release Job\n请使用对应的功能相关动作：暂停工单==>Hold工单；启动工单==>Release工单 ');return false;}else{ var txt=prompt('Please enter the reason to <%=actionalt(j)%> job.请输入<%=actioncalt(j)%>的理由！','');if(txt!=null&&txt!=''){window.open('<%=actionpage(j)%>.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>&reason='+txt,'_self')}}">
        <%next
		  if rs("CONTROL_TYPE")<>"" then%>
        <span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobActionsList.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>')">H</span>
        <%end if
		  set rsR=server.CreateObject("adodb.recordset")
		  SQLR="select * from JOB_ACTIONS_REPEATED where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
		  rsR.open SQLR,conn,1,3
			  if not rsR.eof then%>
        <span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobStationRepeatedList.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>')">R</span>
        <%
			  end if
		  rsR.close
		  set rsR=nothing
		  if DBA=true then
			  if stations_index<>part_stations_index then%>
        <span class="red" style="cursor:hand" title="Update Routine" onClick="window.open('JobPartUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
        <%end if
		  end if%>
        &nbsp;</div></td>
    <%if DBA=true then%>
    <td><div align="center"><img style="cursor:hand" src="/Images/Delete.gif" width="10" height="10" align="absmiddle" onClick="javascript:window.open('DeleteJob.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')"></div></td>
    <%end if
	  end if%>
    <td height="20"><div align="center" class="<%if rs("JOB_TYPE")="N" then%>LinkBlue<%else%>LinkGreen<%end if%>"><strong><span style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>','_blank')"><%=rs("JOB_NUMBER")%>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></span></strong></div></td>
    <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
    <td><div align="center"><%=rs("PART_NUMBER")%></div></td>
    <td><div align="center"><%=rs("LINE_NAME")%></div></td>
    <td><div align="center"><a href='<%=application("KEB_BPS_System_DoNet")%>web/JOB2DCODE/JOB2DCODE.aspx?language=<%=session("language")%>&jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>' target="_blank"><%=rs("JOB_START_QUANTITY")%></a>
        <%if DBA=true then%>
        <span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('JobQuantityUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
        <%end if%>
      </div></td>
    <td height="20"><div align="center">
        <% =formatdate(rs("START_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
    <td><div align="center">
        <% =formatdate(rs("CLOSE_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
    <td><div align="center">
        <%if rs("CYCLE_TIME")<>"" then%>
        <% =formatnumber(clng(rs("CYCLE_TIME"))/60,2,-1)%>
        h
        <%else%>
        &nbsp;
        <%end if%>
      </div></td>
    <%if DisplayStyle="Station" or DisplayStyle="" then%>
    <td height="20"><div align="left">
        <%
		  stations_name_index=getStation(true,"TEXT",""," where S.NID in('"&replace(stations_index,",","','")&"')","",stations_index," -> ")
		  transaction_change=getStationTransactionChange(" where NID in('"&replace(stations_index,",","','")&"')",stations_index)
		  astation=split(stations_index,",")
		  astation_name=split(stations_name_index," -> ")
		  if isnull(transaction_type) or transaction_type="" then
			for j=0 to ubound(astation)
			transaction_type=transaction_type&"0,"
			next
			transaction_type=left(transaction_type,len(transaction_type)-1)
		  end if
		  atransaction_type=split(transaction_type,",")
		  atransaction_change=split(transaction_change,",")
		  if ubound(atransaction_type)<>ubound(astation) then
		  org=ubound(astation)
		  diff=ubound(astation)-ubound(atransaction_type)
		  ReDim Preserve atransaction_type(UBound(atransaction_type) +diff)
		  end if
			 
		  if repeated_stations_sequence<>"" then
		  repeated_stations_sequence=left(repeated_stations_sequence,len(repeated_stations_sequence)-1)
		  end if
			
		 dim CurrentStation
		
		  for j=0 to ubound(astation)
			  if rs("STATUS")="1" then 'job is finished
				if finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then
					showStation=astation(j)
					if ubound(astation_name) >=j then
						showStation=astation_name(j)
					 end if
					station_text=station_text&"<span class='green'>"&showStation&"</span> -> "
				else
					if atransaction_type(j)="0" then
						station_text=station_text&"<span class='blue'>"&astation_name(j)&"</span> -> "
					else
						station_text=station_text&"<span class='blue-italic'>"&astation_name(j)&"</span> -> "
					end if
				end if
			  else
				  if isnull(finished_stations_id) = true then
				  	finished_stations_id=""
				  end if
				  if instr(finished_stations_id,astation(j))=0 and instr(opened_stations_id,astation(j))>0 then 'station is current station
					if atransaction_type(j)="2" then
						station_text=station_text&"<span class='red'>["&astation_name(j)&"]</span> -> "
						CurrentStation=astation_name(j)
					else
						station_text=station_text&"<span class='red'>"&astation_name(j)&"</span> -> "
						CurrentStation=astation_name(j)
					end if
				  elseif finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then 'station is finished
					'if atransaction_type(j)="0" then
					station_text=station_text&"<span class='green'>"&astation_name(j)&"</span> -> "
					'else
					'station_text=station_text&"<span class='green-italic'>"&astation_name(j)&"</span> -> "
					'end if
				  else
					if atransaction_change(j)="0" and atransaction_type(j)<>"2" then
					station_text=station_text&"<span class='blue'>"&astation_name(j)&"</span> -> "
					elseif atransaction_change(j)="0" and atransaction_type(j)="2" then
					station_text=station_text&"<span class='blue'>["&astation_name(j)&"]</span> -> "
					else
						if atransaction_type(j)="0" then
						station_change="1" 'going to change to be optional
						style_class="blue"
						style_word="optional"
						else
						station_change="0"
						style_class="blue-italic" 'going to change to be compulsory
						style_word="compulsory"
						end if
					station_text=station_text&astation_name(j)&" -> "
					end if
				  end if
			  end if
		  next%>
        <%=left(station_text,len(station_text)-3)%></div></td>
    <%
		else
			  stations_name_index=getStation(true,"TEXT",""," where S.NID in('"&replace(stations_index,",","','")&"')","",stations_index," -> ")
			  transaction_change=getStationTransactionChange(" where NID in('"&replace(stations_index,",","','")&"')",stations_index)
			  astation=split(stations_index,",")
			  astation_name=split(stations_name_index," -> ")
			  if isnull(transaction_type) or transaction_type="" then
				for j=0 to ubound(astation)
				transaction_type=transaction_type&"0,"
				next
				transaction_type=left(transaction_type,len(transaction_type)-1)
			  end if
			  atransaction_type=split(transaction_type,",")
			  atransaction_change=split(transaction_change,",")
			  if ubound(atransaction_type)<>ubound(astation) then
			  org=ubound(astation)
			  diff=ubound(astation)-ubound(atransaction_type)
			  ReDim Preserve atransaction_type(UBound(atransaction_type) +diff)
			  end if
				 
			  if repeated_stations_sequence<>"" then
			  repeated_stations_sequence=left(repeated_stations_sequence,len(repeated_stations_sequence)-1)
			  end if
				
			 dim CurrentStation2
			 dim LastStationGroup
			 
			  for j=0 to ubound(astation)
				  CurrentStationGroup=GetStationGroupName(astation(j),rs("JOB_NUMBER"),rs("SHEET_NUMBER"))
				  
				  if CurrentStationGroup<> LastStationGroup then
					  if rs("STATUS")="1" then 'job is finished
						if finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then
							station_text=station_text&"<span class='green'>"&CurrentStationGroup&"</span> -> "
						else
							if atransaction_type(j)="0" then
							station_text=station_text&"<span class='blue'>"&CurrentStationGroup&"</span> -> "
							else
							station_text=station_text&"<span class='blue-italic'>"&CurrentStationGroup&"</span> -> "
							end if
						end if
					  else
					    if isnull(finished_stations_id) = true then
							finished_stations_id=""
					    end if
				  		if instr(finished_stations_id,astation(j))=0 and instr(opened_stations_id,astation(j))>0 then 'station is current station
							if atransaction_type(j)="2" then
							station_text=station_text&"<span class='red'>["&CurrentStationGroup&"]</span> -> "
							CurrentStation2=CurrentStationGroup
							else
							station_text=station_text&"<span class='red'>"&CurrentStationGroup&"</span> -> "
							CurrentStation2=CurrentStationGroup
							end if
						  elseif finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then 'station is finished
							'if atransaction_type(j)="0" then
							station_text=station_text&"<span class='green'>"&CurrentStationGroup&"</span> -> "
							'else
							'station_text=station_text&"<span class='green-italic'>"&astation_name(j)&"</span> -> "
							'end if
						  else
							if atransaction_change(j)="0" and atransaction_type(j)<>"2" then
							station_text=station_text&"<span class='blue'>"&CurrentStationGroup&"</span> -> "
							elseif atransaction_change(j)="0" and atransaction_type(j)="2" then
							station_text=station_text&"<span class='blue'>["&CurrentStationGroup&"]</span> -> "
							else
								if atransaction_type(j)="0" then
								station_change="1" 'going to change to be optional
								style_class="blue"
								style_word="optional"
								else
								station_change="0"
								style_class="blue-italic" 'going to change to be compulsory
								style_word="compulsory"
								end if
							station_text=station_text&CurrentStationGroup&" -> "
							end if
						  end if
					  end if
				  else
				  		if current_station_id=astation(j) then 'station is current station
							if atransaction_type(j)="2" then
								station_text=replace(station_text,CurrentStationGroup,"<span class='red'>["&CurrentStationGroup&"]</span> -> ")
								CurrentStation2=CurrentStationGroup
							else
								station_text=replace(station_text,CurrentStationGroup,"<span class='red'>"&CurrentStationGroup&"</span>")
								CurrentStation2=CurrentStationGroup
							end if
						end if 
				  end if 
					LastStationGroup=CurrentStationGroup
			  next%>
    <td><%=left(station_text,len(station_text)-3)%>
      </div></td>
    </td>
    <%end if %>
	<!--
    <td><% 'add by Jack Zhang 2009-7-17
	  	 if CurrentStation2="" then
	  		response.write "&nbsp;"
		else
	  		response.write(CurrentStation2)
		end if%>
    </td>
	-->
    <td><div align="center"><%=getStationOperator(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("JOB_TYPE"),replace(stations_index,",","','"),repeated_stations_sequence,stations_index," -> ")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("JOB_DEFECTCODE_QUANTITY")%></div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center"><span id="inner_Records"></span></div></td>
  </tr>
  <%end if
rs.close%>
  <tr>
    <td height="20" colspan="16"><!--#include virtual="/Components/JobNote.asp" --></td>
  </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
