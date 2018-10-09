<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobStation.asp" -->
<!--#include virtual="/Functions/GetJobStationMaterialConsume.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
SQL="select J.*,P.PART_NUMBER,P.STATIONS_ROUTINE,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and J.JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	part_number_tag=rs("PART_NUMBER_TAG")
	part_number=rs("PART_NUMBER")
	part_number_id=rs("PART_NUMBER_ID")
	line_name=rs("LINE_NAME")
	job_status=rs("STATUS")
	job_start_time=rs("START_TIME")
	job_close_time=rs("CLOSE_TIME")
	job_start_quantity=rs("JOB_START_QUANTITY")
	job_good_quantity=rs("JOB_GOOD_QUANTITY")
	job_defectcode_quantity=rs("JOB_DEFECTCODE_QUANTITY")
	cycle_time=rs("CYCLE_TIME")
	factory=rs("FACTORY_ID")
	current_station_id=rs("CURRENT_STATION_ID")
		if rs("OPENED_STATIONS_ID")<>"" then
		opened_stations_id=left(rs("OPENED_STATIONS_ID"),len(rs("OPENED_STATIONS_ID"))-1)
		else
		opened_stations_id=""
		end if
		if rs("FINISHED_STATIONS_ID")<>"" then
		finished_stations_id=left(rs("FINISHED_STATIONS_ID"),len(rs("FINISHED_STATIONS_ID"))-1)
		else
		finished_stations_id=""
		end if
		if isnull(rs("STATIONS_INDEX")) or rs("STATIONS_INDEX")="" then
		stations_index=rs("P_STATIONS_INDEX")
		stations_transaction=rs("P_STATIONS_TRANSACTION")
		else
		stations_index=rs("STATIONS_INDEX")
		stations_transaction=rs("STATIONS_TRANSACTION")
		end if
		if rs("STATIONS_ROUTINE")="1" and opened_stations_id<>"" then
		stations_index=opened_stations_id
		end if
	control_type=rs("CONTROL_TYPE")
	astation=split(stations_index,",")
	repeated_stations_sequence=""
		if isnull(rs("REPEATED_STATIONS_SEQUENCE")) or rs("REPEATED_STATIONS_SEQUENCE")="" then
			for i=0 to ubound(astation)
				repeated_stations_sequence=repeated_stations_sequence&"1,"
			next
		else
		repeated_stations_sequence=rs("REPEATED_STATIONS_SEQUENCE")
		end if
	repeated_stations_sequence=left(repeated_stations_sequence,len(repeated_stations_sequence)-1)
	arepeated=split(repeated_stations_sequence,",")
		if ubound(astation)<>ubound(arepeated) then
		diff=ubound(astation)-ubound(arepeated)
		ReDim Preserve arepeated(UBound(arepeated) +diff)
		end if
	if rs("CYCLE_TIME")<>"" then
		if rs("SHIFT_IN_TIME")<>"" and rs("SHIFT_OUT_TIME")<>"" then
			shift_in_time=left(rs("SHIFT_IN_TIME"),len(rs("SHIFT_IN_TIME"))-1)
			a_shift_in_time=split(shift_in_time,",")
			shift_out_time=left(rs("SHIFT_OUT_TIME"),len(rs("SHIFT_OUT_TIME"))-1)
			a_shift_out_time=split(shift_out_time,",")
		end if
		elapsed_time=datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))-shift_interval
	else
		if rs("START_TIME")<>"" and rs("CLOSE_TIME")<>"" then
		elapsed_time=datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))
		else
		elapsed_time=0
		end if
	end if
end if
rs.close

if sheetnumber>=500 and jobtype="R" then
 sql="select a.job_number,b.subjoblist,a.batchno from REWORK_PRINTING a, label_print_history b  where a.batchno= b.batchno and a.JOB_NUMBER='"&jobnumber&"' and a.SEQ='"&sheetnumber&"'"

 rs.open SQL,conn,1,3
	if not rs.eof then
    job_numberold=rs("job_number")
	sheetnumberold=rs("subjoblist")
	batchno=rs("batchno")
    end if
 rs.close

  if  sheetnumberold>500 then
      jobtypeOld="R" 
  else
      jobtypeOld="N"
  end if
end if  




if(request.QueryString("action")="1") then
	SQL="select STATIONS_INDEX,LINE_NAME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		stations_index=rs("STATIONS_INDEX")
		line_name=rs("line_name")
	end if
	rs.close
	'add by jack zhang 
	SQL="delete  JOB_STATION_SCRAP_DETAIL  where job_number='"+jobnumber+"' and sheet_number='"+sheetnumber+"' "
	set rsJOB_STATION_SCRAP_DETAIL=server.CreateObject("adodb.recordset")
	rsJOB_STATION_SCRAP_DETAIL.open SQL,conn,1,3
	
	SQL="delete  JOB_STATION_SCRAP  where job_number='"+jobnumber+"' and sheet_number='"+sheetnumber+"' "
	set rsJOB_STATION_SCRAP=server.CreateObject("adodb.recordset")
	rsJOB_STATION_SCRAP.open SQL,conn,1,3

	arrStationIndexNew=split(stations_index,",")
	for mmmm=0 to ubound(arrStationIndexNew)
		if(arrStationIndexNew(mmmm)<>"") then 	
			errormessage=GetJobStationMaterialConsume(jobnumber,sheetnumber,arrStationIndexNew(mmmm),Line_Name)
		end if 
	next 
end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Functions/TableControl.asp" -->
<!--#include virtual="/Language/Job/SubJobs/Lan_JobDetail.asp" -->
<!--#include virtual="/Language/Components/Lan_JobNote.asp" -->
</head>
<body onLoad="language();language_jobnote()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_Browse"></span> (<%=jobnumber%>-<%=repeatstring(sheetnumber,"0",3)%>) <%if sheetnumber>=500 and jobtype="R" then%>&nbsp;&nbsp;&nbsp;----->>&nbsp;<span id=""  style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/JobDetail.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=repeatstring(sheetnumberold,"0",3)%>&jobtype=<%=jobtypeOld%>')"><%= job_numberold %>-<%=repeatstring(sheetnumberold,"0",3)%><%end if%></span>&nbsp;&nbsp;&nbsp;&nbsp;----->><%=batchno%></td>
  </tr>
  <tr>
    <td height="20" colspan="4" class="t-t-Borrow"><span id="inner_Summary"></span></td>
  </tr>
  <tr>
    <td width="109"><span id="inner_JobNumber"></span></td>
    <td width="229"><%= jobnumber %>-<%=repeatstring(sheetnumber,"0",3)%></td>
    <td width="111"><span id="inner_PartNumber"></span></td>
    <td width="215" height="20"><%= part_number_tag %></td>
  </tr>
  <tr>
    <td><span id="inner_Line"></span></td>
    <td><%= line_name %></td>
    <td><span id="inner_PartType"></span></td>
    <td height="20"><%= part_number %></td>
  </tr>
  <tr>
    <td><span id="inner_Status"></span></td>
    <td height="20"><%
	  if job_status="0" then
	  simg="Opened"
	  aimg="Pause,Abort"
	  alt="pause,abort"
	  apage="PauseJob,AbortJob"
	  elseif job_status="1" then
	  simg="Closed"
	  aimg=""
	  alt=""
	  apage=""
	  elseif job_status="2" then
	  simg="Paused"
	  aimg="Start"
	  alt="Start"
	  apage="StartJob"
	  elseif job_status="3" then
	  simg="Locked"
	  aimg="Start"
	  alt="Unlock"
	  apage="UnlockJob"
	  elseif job_status="4" then
	  simg="Aborted"
	  aimg=""
	  alt=""
	  apage=""
	  elseif job_status="5" then
	  simg="Shift-out"
	  aimg=""
	  alt=""
	  apage=""
	  end if%>
      <img src="/Images/<%=simg%>.gif">
      <%if control_type<>"" then%>
      <input type="button" name="Submit2" value="Actions List" onClick="javascript:window.open('JobActionsList.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>')">
      <%end if%></td>
    <td height="20"><span id="inner_Action"></span></td>
    <td height="20"><%if admin=true then
	  	actionimg=split(aimg,",")
	  	actionalt=split(alt,",")
	  	actionpage=split(apage,",")
	  		for i=0 to ubound(actionimg)%>
      <img style="cursor:hand" src="/Images/<%=actionimg(i)%>.gif" alt="click to <%=actionalt(i)%> job" onClick="javascript:window.open('<%=actionpage(i)%>.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&path=<%=path%>&query=<%=query%>')">
      <%	next
		end if%>
      &nbsp;</td>
  </tr>
  <tr>
    <td><span id="inner_StartTime"></span></td>
    <td height="20"><%= job_start_time %></td>
    <td height="20"><span id="inner_CloseTime"></span></td>
    <td height="20"><%if DBA=true then%>
      <input name="close_time" type="text" id="close_time" value="<%=job_close_time%>" size="20">
      <input type="button" name="Submit3" value="Update" onClick="javascript:location.href='JobCloseTime1.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>&close_time='+document.all.close_time.value+'&path=<%=path%>&query=<%=query%>'">
      <%else%>
      <%= job_close_time %>
      <%end if%>
      &nbsp;</td>
  </tr>
  <tr>
    <td><span id="inner_StartQuantity"></span></td>
    <td height="20"><%= job_start_quantity %></td>
    <td height="20"><span id="inner_GoodQuantity"></span></td>
    <td height="20"><%= job_good_quantity	 %></td>
  </tr>
  <tr>
    <td><span id="inner_DefectQuantity"></span></td>
    <td height="20"><span class="red" style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/DefectCodeDetail.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>')"><%= job_defectcode_quantity %></span>&nbsp; <span id="inner_DefectHistory" class="red" style="cursor:hand" onClick="javascript:window.open('/Job/SubJobs/Defect_Change_History.asp?action=<%=jobnumber&"-"&repeatstring(sheetnumber,"0",3)%>')">dfd</span></td>
    <td height="20"><span id="inner_CycleTime"></span>&nbsp;</td>
    <td height="20"><% =formatnumber(elapsed_time/60,2,-1)%>
      h&nbsp;</td>
  </tr>
</table>
<!--#include virtual="/Components/JobNote.asp" -->
<form name="checkform" method="post" action="JobDetail1.asp">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <Td colspan="20"><input type="button" name="updateFinalYield" value="Update Material Scrap" onclick="checkform.action='jobdetail.asp?action=1&jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>&line_name=<%=line_name%>'; checkform.submit();"></Td>
    </tr>
    <tr>
      <td height="20" colspan="13" class="t-t-Borrow"><span id="inner_StationsInfo"></span></td>
    </tr>
    <tr class="t-t-Borrow">
      <td><div align="center"><span id="inner_NO"></span></div></td>
      <td><div align="center"><span id="inner_Update"></span></div></td>
      <td height="20"><div align="center"><span id="inner_Station"></span></div></td>
      <td><div align="center"><span id="inner_Operator"></span></div></td>
      <td><div align="center"><span id="inner_StationStartTime"></span></div></td>
      <td><div align="center"><span id="inner_StationCloseTime"></span></div></td>
      <td><div align="center"><span id="inner_Elapsed"></span></div></td>
      <td><div align="center"><span id="inner_Minus"></span></div></td>
      <td><div align="center"><span id="inner_ShiftOutTime"></span></div></td>
      <td><div align="center"><span id="inner_ShiftInTime"></span></div></td>
      <td><div align="center"><span id="inner_Interval"></span></div></td>
      <td><div align="center"><span id="inner_Equal"></span></div></td>
      <td><div align="center"><span id="inner_Cycle"></span></div></td>
    </tr>
    <%
  SQL="select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where NID in ('"&replace(stations_index,",","','")&"')"
  rs.open SQL,conn,1,3
  if not rs.eof then
  i=1
  for j=0 to ubound(astation)
	  while not rs.eof
	  	  station_finished=false
		  operatorcode=null
		  practised=null
		  station_start_time=null
		  station_close_time=null
		  elapsed=0
		  this_shift_out_time=null
		  this_shift_in_time=null
		  sinterval=0
		  interval=0
		  cycle=0
		  unit=null		
			if rs("NID")=astation(j) then
				if finished_stations_id<>"" and instr(finished_stations_id,rs("NID"))<>0 then
				unit="m"
				station_finished=true
				getJobStation jobnumber,sheetnumber,jobtype,rs("NID"),arepeated(j),operatorcode,practised,station_start_time,station_close_time'get the finished stations info
					if IsArray(a_shift_out_time) then
					for s=0 to ubound(a_shift_out_time)
						'if cdate(a_shift_out_time(s))>=cdate(before_station_close_time) and cdate(a_shift_in_time(s))<=cdate(station_start_time) then
						'sinterval=datediff("n",a_shift_out_time(s),a_shift_in_time(s))
						%>
    <tr>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20">&nbsp;</td>
      <td height="20"><%=a_shift_out_time(s)%>&nbsp;</td>
      <td height="20"><%=a_shift_in_time(s)%>&nbsp;</td>
      <td height="20"><div align="center"><%=sinterval%>&nbsp;</div></td>
      <td height="20"><div align="center">=</div></td>
      <td height="20"><div align="center"><%=sinterval%>&nbsp;<%=unit%></div></td>
    </tr>
    <%			'end if
		  			next
					end if
				end if%>
    <tr>
      <td><div align="center"><%=i%></div></td>
      <td><div align="center">
          <input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1">
        </div></td>
      <td height="20"><span style="cursor:hand" onClick="tableexpand(<%=i%>,'/Job/SubJobs/StationDetail.asp?jobnumber=<%=jobnumber%>&sheetnumber=<%=sheetnumber%>&jobtype=<%=jobtype%>&station_id=<%=rs("NID")%>&repeated_sequence=<%=arepeated(j)%>&factory=<%=factory%>&part_number=<%=part_number%>&part_number_id=<%=part_number_id%>',20)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=i%>" width="9" height="9" border="0" align="absmiddle" id="tabimg<%=i%>" title="Expand"><span 
	  <%if station_finished=true then'get the finished stations info%>
	  class="StrongGreen"> <%=rs("STATION_NAME")%>&nbsp;<%=rs("STATION_CHINESE_NAME")%>&nbsp;<%=arepeated(j)%></span>
        <%else%>
			<%'if current_station_id=rs("NID") then%>
			<%if isnull(finished_stations_id) = true then
				finished_stations_id=""
		     end if
			 if instr(finished_stations_id,astation(j))=0 and instr(opened_stations_id,astation(j))>0 then
			  %>
			class="strongred"
			<%end if%>
        ><%=rs("STATION_NAME")%>&nbsp;<%=rs("STATION_CHINESE_NAME")%>&nbsp;<%=arepeated(j)%>
        <%end if%>
        <%=astation(j)%>
        <input name="station_id<%=i%>" type="hidden" id="station_id<%=i%>" value="<%=rs("NID")%>">
        <input name="repeated_sequence<%=i%>" type="hidden" id="repeated_sequence<%=i%>" value="<%=arepeated(j)%>">
        </span></td>
      <td <%if practised="1" then%>class="t-c-practised"<%end if%>><div align="center">
          <input name="operator_code<%=i%>" type="text" id="operator_code<%=i%>" value="<%=operatorcode%>" size="4" <%if practised="1" then%>class="t-c-practised"<%end if%>>
        </div></td>
      <td><div align="center">
          <input name="start_time<%=i%>" type="text" id="start_time<%=i%>" value="<%=station_start_time%>">
        </div></td>
      <td><div align="center">
          <input name="stop_time<%=i%>" type="text" id="stop_time<%=i%>" value="<%=station_close_time%>">
        </div></td>
      <td><div align="center">
          <%elapsed=datediff("n",station_start_time,station_close_time)%>
          <%=elapsed%>&nbsp;<%=unit%></div></td>
      <td><div align="center">-</div></td>
      <td><%
	 	if IsArray(a_shift_out_time) then
	 	for s=0 to ubound(a_shift_out_time)
			'if cdate(a_shift_out_time(s))>=cdate(station_start_time) and cdate(a_shift_out_time(s))<=cdate(station_close_time) then
'			this_shift_out_time=a_shift_out_time(s)%>
        <%'=this_shift_out_time%>
        <%
'			end if
		next
		end if%>
        &nbsp;</td>
      <td><%
		if IsArray(a_shift_out_time) then
		for s=0 to ubound(a_shift_in_time)
			'if cdate(a_shift_in_time(s))>=cdate(station_start_time) and cdate(a_shift_in_time(s))<=cdate(station_close_time) then
			'this_shift_in_time=a_shift_in_time(s)%>
        <%'=this_shift_in_time%>
        <%
			'end if
		next
		end if%>
        &nbsp;</td>
      <td><div align="center">
          <%if isnull(this_shift_out_time)=false and isnull(this_shift_in_time)=false then
		interval=datediff("n",this_shift_out_time,this_shift_in_time)
		end if%>
          <%=interval%>&nbsp;</div></td>
      <td><div align="center">=</div></td>
      <td><div align="center">
          <%cycle=elapsed-interval%>
          <%=cycle%>&nbsp;<%=unit%></div></td>
    </tr>
    <tbody id="tab<%=i%>" style="display:none">
      <tr>
        <td>&nbsp;</td>
        <td colspan="12"><iframe src="" width="100%" height="200" scrolling="auto" frameborder="0" id="tabfrm<%=i%>"></iframe></td>
      </tr>
    </tbody>
    <%before_station_close_time=station_close_time
	  	end if
	  rs.movenext
	  wend
  rs.movefirst
  i=i+1
  next
  end if
  rs.close
  set rs=nothing%>
    <tr>
      <td colspan="13"><div align="center">
          <input name="idcount" type="hidden" id="idcount" value="<%=i-1%>">
          <input name="jobnumber" type="hidden" id="jobnumber" value="<%=jobnumber%>">
          <input name="sheetnumber" type="hidden" id="sheetnumber" value="<%=sheetnumber%>">
          <input name="jobtype" type="hidden" id="jobtype" value="<%=jobtype%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input name="CheckAll" type="button" id="CheckAll" value="Check All" onClick="checkall()">
          &nbsp;
          <input name="UncheckAll" type="button" id="UncheckAll" value="Uncheck All" onClick="uncheckall()">
          &nbsp;
          <input name="Update" type="submit" id="Update" value="Updated" <%if admin=false then%>disabled<%end if%>>
          &nbsp;
          <input name="Reset" type="reset" id="Reset" value="Reset">
        </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
