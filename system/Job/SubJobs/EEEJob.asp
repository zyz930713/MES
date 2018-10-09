<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
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
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if jobstatus<>"" then
where=where&" and J.STATUS="&jobstatus
end if
if currentstation<>"" then
where=where&" and J.CURRENT_STATION_ID='"&currentstation&"'"
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
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Job/EEEJob.asp"
SQL="select 1,J.*,P.PART_NUMBER,P.STATIONS_ROUTINE,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J left join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER like '%E'"&where&order
rs.open SQL,conn,1,3
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
</head>

<body>
<form action="/Job/SubJobs/EEEJob.asp" method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"><span>Search Job </span></td>
    </tr>
    <tr>
      <td height="20"><span class="style1">Job Number</span>        </td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
      <td>Part Number </td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
      <td>Status</td>
      <td><select name="jobstatus" id="jobstatus">
        <option value="">All</option>
        <option value="0" <%if jobstatus="0" then%>selected<%end if%>>Opened</option>
        <option value="2" <%if jobstatus="2" then%>selected<%end if%>>Paused</option>
        <option value="1" <%if jobstatus="1" then%>selected<%end if%>>Closed</option>
		<option value="3" <%if jobstatus="3" then%>selected<%end if%>>Locked</option>
      </select></td>
    </tr>
    <tr>
      <td height="20">Current Station </td>
      <td height="20"><select name="currentstation" id="currentstation">
        <option value="">All</option>
        <%=getStation(true,"OPTION",currentstation,"","","","")%>
      </select></td>
      <td>Job Start Time</td>
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
      <td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy">Browse EEE Job list (Default in past 7 days) </td>
    </tr>
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy">User:
      <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
      <td class="t-t-Borrow"><div align="center">Status</div></td>
	  <%if admin=true then%>
      <td height="20" class="t-t-Borrow"><div align="center">Action</div></td>
	  <%if DBA=true then%>
	  <td class="t-t-Borrow"><div align="center">Delete</div></td>
	  <%end if
	  end if%>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Job Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Part Number <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Start Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center">Close Time </div></td>
      <td height="20" class="t-t-Borrow"><div align="center">Included Stations </div></td>
      <td class="t-t-Borrow"><div align="center">Stations' Operators </div></td>
    </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
station_text=""
stations_index=rs("STATIONS_INDEX")
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
stations_index=rs("P_STATIONS_INDEX")
transaction_type=rs("P_STATIONS_TRANSACTION")
end if
if rs("STATIONS_ROUTINE")="1" then
stations_index=opened_stations_id
end if
%>
    <tr>
      <td height="20"><div align="center">
        <% =(session("strpagenum")-1)*pagesize_s+i%>
</div></td>
      <td><div align="center">
	  <%
	  if rs("STATUS")="0" then
	  simg="Opened"
	  aimg="Pause,Abort"
	  alt="pause,abort"
	  apage="PauseJob,AbortJob"
	  elseif rs("STATUS")="1" then
	  simg="Closed"
	  aimg="Repeat"
	  alt="Repeat"
	  apage="RepeatJob"
	  elseif rs("STATUS")="2" then
	  simg="Paused"
	  aimg="Start"
	  alt="Start"
	  apage="StartJob"
	  elseif rs("STATUS")="3" then
	  simg="Locked"
	  aimg="Start"
	  alt="Unlock"
	  apage="UnlockJob"
	  elseif rs("STATUS")="4" then
	  simg="Aborted"
	  aimg=""
	  alt=""
	  apage=""
	  end if%>
	  <img src="/Images/<%=simg%>.gif"></div></td>
      <%if admin=true then%>
	  <td height="20" class="red"><div align="center"><%
	  actionimg=split(aimg,",")
	  actionalt=split(alt,",")
	  actionpage=split(apage,",")
	  for j=0 to ubound(actionimg)%>
	  <img style="cursor:hand" src="/Images/<%=actionimg(j)%>.gif" alt="click to <%=actionalt(j)%> job" onClick="javascript:var txt=prompt('Please enter the reason to <%=actionalt(j)%> job.','');if(txt!=null&&txt!=''){window.open('<%=actionpage(j)%>.asp?jobnumber=<%=rs("JOB_NUMBER")%>&path=<%=path%>&query=<%=query%>&reason='+txt,'_self')}">
	  <%next
	  if rs("CONTROL_TYPE")<>"" then%><span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobActionsList.asp?jobnumber=<%=rs("JOB_NUMBER")%>')">H</span>
	  <%end if
	  set rsR=server.CreateObject("adodb.recordset")
	  SQLR="select * from JOB_ACTIONS_REPEATED where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
	  rsR.open SQLR,conn,1,3
	  if not rsR.eof then%><span class="red" style="cursor:hand" title="Browse actions" onClick="window.open('JobStationRepeatedList.asp?jobnumber=<%=rs("JOB_NUMBER")%>')">R</span>
	  <%end if
	  rsR.close
	  set rsR=nothing%>
	  &nbsp;</div></td>
	  <%if DBA=true then%>
	  <td><div align="center"><img style="cursor:hand" src="/Images/Delete.gif" width="10" height="10" align="absmiddle" onClick="javascript:window.open('DeleteJob.asp?jobnumber=<%=rs("JOB_NUMBER")%>&path=<%=path%>&query=<%=query%>','_self')"></div></td>
	  <%end if
	  end if%>
      <td height="20"><div align="center" class="d_link"><strong><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%=rs("JOB_NUMBER")%></a></strong></div></td>
      <td><div align="center"><%=rs("PART_NUMBER")%></div></td>
      <td height="20"><div align="center"><% =formatdate(rs("START_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
      <td><div align="center"><% =formatdate(rs("CLOSE_TIME"),application("longdateformat"))%>&nbsp;</div></td>
      <td height="20"><div align="left">
	  <%
	  if stations_index<>"" then
		  stations_name_index=getStation(true,"TEXT",""," where NID in('"&replace(stations_index,",","','")&"')","",stations_index," -> ")
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
			  if rs("STATUS")="1" then 'station is finished
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
					'if atransaction_type(j)="0" then
					station_text=station_text&"<span class='red'>"&astation_name(j)&"</span> -> "
					'else
					'station_text=station_text&"<span class='red-italic'>"&astation_name(j)&"</span> -> "
					'end if
				  elseif finished_stations_id<>"" and instr(finished_stations_id,astation(j))<>0 then 'station is finished
					'if atransaction_type(j)="0" then
					station_text=station_text&"<span class='green'>"&astation_name(j)&"</span> -> "
					'else
					'station_text=station_text&"<span class='green-italic'>"&astation_name(j)&"</span> -> "
					'end if
				  else
					if atransaction_change(j)="0" then
					station_text=station_text&"<span class='blue'>"&astation_name(j)&"</span> -> "
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
					station_text=station_text&"<span class='"&style_class&"' style='cursor:hand' title='Click this station to be "&style_word&".' onClick=""javascript:location.href='JobStationChange.asp?jobnumber="&rs("JOB_NUMBER")&"&station_id="&astation(j)&"&type="&station_change&"&path="&path&"&query="&query&"'"">"&astation_name(j)&"</span> -> "
					end if
				  end if
			  end if
		  next
	  %>
	  <%=left(station_text,len(station_text)-3)%>
	  <%end if%>&nbsp;</div></td>
      <td><div align="left"><%if stations_index<>"" then%><%=getStationOperator(rs("JOB_NUMBER"),replace(stations_index,",","','"),repeated_stations_sequence,stations_index," -> ")%><%end if%>&nbsp;</div></td>
    </tr>
<%
i=i+1
rs.movenext
wend
else
%>
    <tr>
      <td height="20" colspan="10"><div align="center">No Records</div></td>
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
