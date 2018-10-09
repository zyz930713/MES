<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
sheetnumber=trim(request("sheetnumber"))
jobtype=trim(request("jobtype"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
factory=trim(request("factory"))
timespan=trim(request("timespan"))
jobstatus=trim(request("jobstatus"))
currentstation=trim(request("currentstation"))
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
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if sheetnumber<>"" then
where=where&" and J.SHEET_NUMBER='"&sheetnumber&"'"
end if
if jobtype<>"" then
where=where&" and J.JOB_TYPE='"&jobtype&"'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME)='"&lcase(line)&"'"
end if
if factory<>"" then
where=where&" and P.FACTORY_ID='"&factory&"'"
end if
if jobstatus<>"" then
where=where&" and J.STATUS="&jobstatus
end if
if currentstation<>"" then
where=where&" and J.CURRENT_STATION_ID='"&currentstation&"'"
end if
if timespan<>"" then
where=where&" and J.CYCLE_TIME>="&timespan*60
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
elseif jobnumber="" then
fromdate=dateadd("d",-7,date())
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&sheetnumber="&sheetnumber&"&jobtype="&jobtype&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&timespan="&timespan&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/System/SubJobRecycle/SubJobRecycle.asp"
SQL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB_BACKUP J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null"&where&order
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
<!--#include virtual="/Language/System/SubJobRecycle/Lan_SubJobRecycle.asp" -->
</head>

<body onLoad="language();language_page();language_jobnote()">
<form action="/System/SubJobRecycle/SubJobRecycle.asp" method="get" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-b-midautumn"><span id="inner_Search"></span></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchJobNumber"></span></td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">
        -
          <input name="sheetnumber" type="text" id="sheetnumber" value="<%=sheetnumber%>" size="3" maxlength="3"></td>
      <td><span id="inner_SearchJobType"></span></td>
      <td><select name="jobtype" id="jobtype">
        <option value="N" <%if jobtype="N" then%>selected<%end if%>>Normal</option>
        <option value="R" <%if jobtype="R" then%>selected<%end if%>>Rework</option>
      </select>
      </td>
      <td><span id="inner_SearchPartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchLineName"></span></td>
      <td height="20"><input name="line" type="text" id="line" value="<%=line%>"></td>
      <td><span id="inner_SearchFactory"></span></td>
      <td><select name="factory" id="factory">
        <option value="">Factory</option>
        <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
        <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
        <%= getFactory("OPTION",factory,null,null,null) %>
      </select></td>
      <td><span id="inner_SearchStatus"></span></td>
      <td><select name="jobstatus" id="jobstatus">
        <option value="">All</option>
        <option value="0" <%if jobstatus="0" then%>selected<%end if%>>Opened/正开放</option>
        <option value="2" <%if jobstatus="2" then%>selected<%end if%>>Paused/暂停</option>
        <option value="1" <%if jobstatus="1" then%>selected<%end if%>>Closed/已结束</option>
        <option value="3" <%if jobstatus="3" then%>selected<%end if%>>Locked/被锁定</option>
		<option value="5" <%if jobstatus="5" then%>selected<%end if%>>Locked/被停线</option>
      </select></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchCurrentStation"></span></td>
      <td height="20"><select name="currentstation" id="currentstation">
        <option value="">All</option>
        <%=getStation(true,"OPTION",currentstation,"","","","")%>
      </select></td>
      <td><span id="inner_SearchCycleTime"></span></td>
      <td>&gt;=
        <input name="timespan" type="text" id="timespan" value="<%=timespan%>" size="4">
        <span id="inner_SearchM"></span>h</td>
      <td><span id="inner_SearchJobStartTime"></span></td>
      <td><span id="inner_SearchFrom"></span><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
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
&nbsp;<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="15" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="48%"><span id="inner_Browse"></span></td>
            <td width="52%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Job_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td height="20" colspan="15" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20" colspan="15"><!--#include virtual="/Components/PageSplit.asp" --></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
	  <td class="t-t-Borrow"><div align="center"><span id="inner_Delete"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_PartType"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_StartTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_CloseTime"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Operators"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_LineLost"></span></div></td>
    </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
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
        <% =(csng(session("strpagenum"))-1)*pagesize_s+i%>
</div></td>
      <td><div align="center">
	  <%SubJobStatus rs("STATUS"),simg,aimg,text,ctext,alt,calt,apage%>
      <img src="/Images/<%=simg%>.gif"></div></td>
	  <td height="20" class="red"><div align="center"><%
	  actionimg=split(aimg,",")
	  actionalt=split(alt,",")
	  actioncalt=split(calt,",")
	  actionpage=split(apage,",")
		  for j=0 to ubound(actionimg)%>
		  <img style="cursor:hand" src="/Images/<%=actionimg(j)%>.gif" alt="click to <%=actionalt(j)%> job.点击<%=actioncalt(j)%>工单" onClick="javascript:var txt=prompt('Please enter the reason to <%=actionalt(j)%> job.请输入<%=actioncalt(j)%>的理由！','');if(txt!=null&&txt!=''){window.open('<%=actionpage(j)%>.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>&reason='+txt,'_self')}">
		  <%next
		  if rs("CONTROL_TYPE")<>"" then%><span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobActionsList.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>')">H</span>
		  <%end if
		  set rsR=server.CreateObject("adodb.recordset")
		  SQLR="select * from JOB_ACTIONS_REPEATED where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
		  rsR.open SQLR,conn,1,3
			  if not rsR.eof then%><span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobStationRepeatedList.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>')">R</span>
			  <%
			  end if
		  rsR.close
		  set rsR=nothing
		  if DBA=true then
			  if stations_index<>part_stations_index then%><span class="red" style="cursor:hand" title="Update Routine" onClick="window.open('JobPartUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
			  <%end if
		  end if%>
	  	&nbsp;</div></td>
		  <td><div align="center"><img style="cursor:hand" src="/Images/Delete.gif" width="10" height="10" align="absmiddle" onClick="javascript:window.open('FullDeleteJob.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')"></div></td>
      <td height="20"><div align="center" class="<%if rs("JOB_TYPE")="N" then%>LinkBlue<%else%>LinkGreen<%end if%>"><strong><span style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>','_blank')"><%=rs("JOB_NUMBER")%>-<%=replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)%></span></strong></div></td>
      <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
      <td><div align="center"><%=rs("PART_NUMBER")%></div></td>
      <td><div align="center"><%=rs("LINE_NAME")%></div></td>
      <td><div align="center"><%=rs("JOB_START_QUANTITY")%><%if DBA=true then%><span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('JobQuantityUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&jobtype=<%=rs("JOB_TYPE")%>&path=<%=path%>&query=<%=query%>','_self')">U</span><%end if%></div></td>
      <td height="20"><div align="center">
        <% =formatdate(rs("START_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
      <td><div align="center"><% =formatdate(rs("CLOSE_TIME"),application("longdateformat"))%>&nbsp;</div></td>
      <td><div align="center">
        <%if rs("CYCLE_TIME")<>"" then%>
		<% =formatnumber(clng(rs("CYCLE_TIME"))/60,2,-1)%> h
		<%else%>&nbsp;
		<%end if%>
      </div></td>
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

	  for j=0 to ubound(astation)
	  	  if rs("STATUS")="1" then 'job is finished
		  	if finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then
		  	station_text=station_text&"<span class='green'>"&astation_name(j)&"</span> -> "
			else
				if atransaction_type(j)="0" then
				station_text=station_text&"<span class='blue'>"&astation_name(j)&"</span> -> "
				else
				station_text=station_text&"<span class='blue-italic'>"&astation_name(j)&"</span> -> "
				end if
			end if
		  else
			  if current_station_id=astation(j) then 'station is current station
			  	if atransaction_type(j)="2" then
				station_text=station_text&"<span class='red'>["&astation_name(j)&"]</span> -> "
				else
			  	station_text=station_text&"<span class='red'>"&astation_name(j)&"</span> -> "
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
				station_text=station_text&"<span class='"&style_class&"' style='cursor:hand' title='Click this station to be "&style_word&".' onClick=""javascript:location.href='JobStationChange.asp?jobnumber="&rs("JOB_NUMBER")&"&sheetnumber="&rs("SHEET_NUMBER")&"&jobtype="&rs("JOB_TYPE")&"&station_id="&astation(j)&"&type="&station_change&"&path="&path&"&query="&query&"'"">"&astation_name(j)&"</span> -> "
				end if
			  end if
		  end if
	  next%>
	  <%=left(station_text,len(station_text)-3)%></div></td>
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
      <td height="20" colspan="15"><div align="center"><span id="inner_Records"></span>No Records</div></td>
    </tr>
<%end if
rs.close%>
	<tr>
      <td height="20" colspan="15"><!--#include virtual="/Components/JobNote.asp" --></td>
    </tr>
</table>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
