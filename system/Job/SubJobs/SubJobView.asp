<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<%
if request("language")="1" then
	session("language")=0
else
	session("language")=1
end if	
fromdate=request("fromdate")
line=getLine()

order=" order by J.JOB_NUMBER desc,J.SHEET_NUMBER desc"
if line<>"" then
	where=where&" and J.LINE_NAME in ('"&line&"')"
end if
if fromdate<>"" then
	'where=where&" and J.START_TIME>=to_date('"&fromdate& "','yyyy-mm-dd')"
	where=where&" and J.START_TIME>=sysdate-0.5"
end if
order=" order by J.START_TIME desc"

SQL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID LEFT JOIN STATION S ON J.CURRENT_STATION_ID=S.NID where 1=1 "&where&order

rs.open SQL,conn,1,3
'response.Write(SQL)

function getLine()
	lineName=""
	SQL="select line_name from computer_printer_mapping where computer_name='"&Request.Cookies("computer_name")&"'"
	set rsLine = server.createobject("adodb.recordset")
	rsLine.open SQL,conn,1,3
	if not rsLine.eof then
		if rsLine("line_name")<>"" then
			lineName=replace(rsLine("line_name"),",","','")
		end if
	end if
	rsLine.close
	getLine=lineName
end function
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<style>
body {
	font-size: 11pt;
	font-family: Verdana, Arial, Helvetica, sans-serif;	
}
table {
	font-size: 10pt;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
span {
	font-size: 12px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
</style>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<!--#include virtual="/Language/Job/SubJobs/Lan_Job.asp" -->
<!--#include virtual="/Language/Components/Lan_JobNote.asp" -->
</head>
<body onLoad="language();language_jobnote()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr style="font-size:12px"><td colspan="12"><table width="100%"><tr><td ><span id="inner_JobProductRecord"></span>&nbsp;(<%=fromdate%>)</td>
  <td align="right" >
  <a href="SubJobView.asp?fromdate=<%=fromdate%>&language=<%=session("language")%>"><%if session("language")=1 then%>English<%else%>¼òÌåÖÐÎÄ<%end if%></a></td></tr></table>
  </td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Status"></span></div></td>    
    <td height="20" class="t-t-Borrow"><div align="center">
	<span id="inner_JobNumber"></span>
	</div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>    
    <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center">	
	<span id="inner_StartTime"></span>
	</div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_CloseTime"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Stations"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Operators"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_LineLost"></span></div></td>
  </tr>
  <%
i=1
if not rs.eof then
while not rs.eof 
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
			SQL="SELECT 1 FROM job_stations where JOB_NUMBER='"+rs("JOB_NUMBER")+"' AND SHEET_NUMBER='"+cstr(rs("SHEET_NUMBER"))+"' and station_id='"+mCurrent_Status_ID+"'"
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
    
    <td height="20"><div align="center" ><%=rs("JOB_NUMBER")%>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></span></div></td>
    <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
    <td><div align="center"><%=rs("LINE_NAME")%></div></td>
    <td><div align="center"><%=rs("JOB_START_QUANTITY")%></div></td>
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
    <td height="20" colspan="16" class="t-c-GrayLight">
		<table border="0" cellpadding="0" cellspacing="1" >
		  <tr class="t-c-GrayLight" align="right">
			<td height="20" ><img src="/Images/Opened.gif" width="10" height="10" align="absmiddle"></td>
			<td>=</td>
			<td height="20"><span id="inner_Jobisopened"></span></td>
			<td>&nbsp;</td>
			<td height="20"><img src="/Images/Paused.gif" width="10" height="10" align="absmiddle"></td>
			<td>=</td>
			<td height="20"><span id="inner_Jobispaused"></span></td>
			<td>&nbsp;</td>
			<td height="20"><img src="/Images/Closed.gif" width="10" height="10"></td>
			<td>=</td>
			<td height="20"><span id="inner_Jobisclosed"></span></td>
			<td>&nbsp;</td>
			<td><img src="/Images/ShiftOut.gif" width="10" height="10" /></td>
			<td>=</td>
			<td><span id="inner_Jobisshiftouted"></span></td>
			<td>&nbsp;</td>
			<td><img src="/Images/Locked.gif" width="10" height="10"></td>
			<td>=</td>
			<td><span id="inner_Jobislocked"></span></td>
			<td>&nbsp;</td>
			<td><img src="/Images/Aborted.gif" width="10" height="10" /></td>
			<td>=</td>
			<td><span id="inner_Jobisaborted"></span></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>			
		  </tr>		
		  <tr>
			<td nowrap="nowrap" class="green"><span id="inner_Green"></span></td>
			<td nowrap="nowrap">=</td>
			<td nowrap="nowrap"><span id="inner_Finishedstation"></span></td>
			<td nowrap="nowrap">&nbsp;</td>
			<td nowrap="nowrap" class="red"><span id="inner_Red"></span></td>
			<td nowrap="nowrap">=</td>
			<td nowrap="nowrap"><span id="inner_Currentstation"></span></td>
			<td nowrap="nowrap">&nbsp;</td>
			<td nowrap="nowrap" class="blue"><span id="inner_Blue"></span></td>
			<td nowrap="nowrap">=</td>
			<td nowrap="nowrap"><span id="inner_Waitingstation"></span></td>
			<td nowrap="nowrap" class="green-italic">&nbsp;</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
