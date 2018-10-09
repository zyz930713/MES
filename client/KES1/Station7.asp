<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/TrackJobMaster.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/GetJobStationMaterialConsume.asp" -->

<%
pagename="Station7.asp"

current_station_id=session("current_station_id")
job_is_closed=false
new_rework=false
next_station=""
next_station_name=""
total_defect_quantity=0

'save all defectcodes' info.
'only normal job has defectcode, rework job does not have.
if (instr("N,R,S,C",session("JOB_TYPE")))>=0 then
	for i=0 to cint(request("defect_count"))
		deftQty=request("defect_quantity"&i)
		if deftQty<>"" and deftQty<>"0" then
			total_defect_quantity = total_defect_quantity + clng(deftQty)
'			SQL="select JOB_NUMBER,SHEET_NUMBER,STATION_ID,DEFECT_CODE_ID,REPEATED_SEQUENCE,DEFECT_QUANTITY from JOB_DEFECTCODES "
'			SQL=SQL+"where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and STATION_ID='"&request("station_id"&i)&"' and DEFECT_CODE_ID='"&request("defect_code"&i)&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
'			rs.open SQL,conn,1,3
'			if rs.eof then
'				rs.addnew
'				rs("JOB_NUMBER")=session("JOB_NUMBER")
'				rs("SHEET_NUMBER")=session("SHEET_NUMBER")
'				rs("STATION_ID")=request("station_id"&i)
'				rs("DEFECT_CODE_ID")=request("defect_code"&i)
'				rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
'				rs("DEFECT_QUANTITY")=deftQty
'			else
'				rs("DEFECT_QUANTITY")=clng(rs("DEFECT_QUANTITY"))+clng(deftQty)
'			end if			
'			rs.update
'			rs.close

			'update job stations info
			SQL = "select start_time,good_quantity,station_start_quantity,station_defectcode_quantity,defectcode_update from job_stations "
			SQL = SQL+"where job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' and job_type='"&session("JOB_TYPE")&"' and station_id='"&request("station_id"&i)&"' and repeated_sequence='"&session("REPEATED_SEQUENCE")&"'"
			rs.open SQL,conn,1,3
			startTime=""
			if not rs.eof then
				startTime = rs("start_time")
				rs("station_defectcode_quantity")=clng(rs("station_defectcode_quantity"))+clng(deftQty)
				rs("good_quantity")=clng(rs("good_quantity"))-clng(deftQty)
				rs("defectcode_update")=1
				rs.update							
			end if			
			rs.close
			
			'update stations which is opened after startTime
			stationList=""
			if startTime <> "" then
				SQL = "select good_quantity,station_start_quantity,station_id from job_stations "
				SQL = SQL + "where start_time>to_date('"&startTime&"','yyyy-mm-dd hh24:mi:ss') and job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' and job_type='"&session("JOB_TYPE")&"' and repeated_sequence='"&session("REPEATED_SEQUENCE")&"'"
				rs.open SQL,conn,1,3
				while not rs.eof 
					rs("good_quantity")=clng(rs("good_quantity"))-clng(deftQty)
					rs("station_start_quantity")=clng(rs("station_start_quantity"))-clng(deftQty)
					stationList=stationList&rs("station_id")&"','"
					rs.update
					rs.movenext
				wend
				rs.close
			end if
			
			'update Job Quantity in JOB_ACTIONS
			if stationList <> "" then
				SQL = "select action_value from job_actions ja where exists (select 1 from action where nid=ja.action_id and action_purpose=5) and ja.station_id in ('"&stationList&"')"
				SQL = SQL +  " and job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' and job_type='"&session("JOB_TYPE")&"' and repeated_sequence='"&session("REPEATED_SEQUENCE")&"'"
				rs.open SQL,conn,1,3
				while not rs.eof 
					rs("action_value")=clng(rs("action_value"))-clng(deftQty)
					rs.update
					rs.movenext
				wend
				rs.close
			end if
		end if
	next	
end if
'delete temp defect
'SQL="delete from job_defectcodes_temp where enter_station_id='"&current_station_id&"'"
'rs.open SQL,conn,1,3

'get job's properties.
SQL="select * from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='" &session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	'to transact job's normal routine
	part_number_id=rs("PART_NUMBER_ID")
	part_number_tag=rs("PART_NUMBER_TAG")
	factory_id=rs("FACTORY_ID")
	line_name=rs("LINE_NAME")
	linename=rs("LINE_NAME")

	stations_routine=rs("STATIONS_ROUTINE")
	set rsC=server.CreateObject("adodb.recordset") 'get stations_index and stations_transaction properties for this job
	if isnull(rs("STATIONS_INDEX")) or rs("STATIONS_INDEX")="" then
		SQLC="select STATIONS_INDEX,STATIONS_TRANSACTION from PART where NID='"&part_number_id&"'"
		rsC.open SQLC,conn,1,3
		if not rsC.eof then
			stations_index=rsC("STATIONS_INDEX")
			stations_transaction=rsC("STATIONS_TRANSACTION")
		end if
		rsC.close
	else
	stations_index=rs("STATIONS_INDEX")
	stations_transaction=rs("STATIONS_TRANSACTION")
	end if
	'close current station
	close_time=now()
	SQLC="update JOB_STATIONS set STATUS=2,CLOSE_TIME='"&close_time&"' where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='" &session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
	rsC.open SQLC,conn,1,3
	
	if stations_routine="0" then'fixed routine
		if current_station_id=rs("LAST_STATION_ID") then'if current station is last station, close this job.
		    job_is_closed=true
		    rs("STATUS")=1
		    rs("CLOSE_TIME")=request.Form("station_close_time")
			'caculate cycle time
			shift_interval=0
			elapsed_time=0
			if rs("SHIFT_IN_TIME")<>"" and rs("SHIFT_OUT_TIME")<>"" then
			    shift_in_time=left(rs("SHIFT_IN_TIME"),len(rs("SHIFT_IN_TIME"))-1)
			    a_shift_in_time=split(shift_in_time,",")
			    shift_out_time=left(rs("SHIFT_OUT_TIME"),len(rs("SHIFT_OUT_TIME"))-1)
			    a_shift_out_time=split(shift_out_time,",")
				if ubound(a_shift_in_time)=ubound(a_shift_out_time) then
					for i=0 to ubound(a_shift_in_time)
					    shift_interval=shift_interval+datediff("n",a_shift_out_time(i),a_shift_in_time(i))
					next					
				end if
			end if
			elapsed_time=cstr(datediff("n",rs("START_TIME"),request.Form("station_close_time"))-shift_interval)
			rs("CYCLE_TIME")=elapsed_time
		else 'current station is not last station
			job_is_closed=false
			'get stations' index
			astation=split(STATIONS_INDEX,",")
			atransaction=split(STATIONS_TRANSACTION,",")
								
			for i=0 to ubound(astation)
				if current_station_id=astation(i) then
					current_station_transaction=atransaction(i)
					for j=i+1 to ubound(astation)
						if atransaction(j)="0" or atransaction(j)="2" then
						next_station=astation(j)
						rs("CURRENT_STATION_ID")=astation(j)'update current station to be next.
						exit for
						end if
					next
					exit for
				end if
			next
			'get station name of next station
			set rsS=server.CreateObject("adodb.recordset")
			SQLS="select STATION_NAME,STATION_CHINESE_NAME from STATION where NID='"&next_station&"'"
			rsS.open SQLS,conn,1,3
			if not rsS.eof then
				next_station_name=rsS("STATION_NAME")
				next_station_chinese_name=rsS("STATION_CHINESE_NAME")
			end if
			rsS.close
			set rsS=nothing
		end if
	else'others routine
		total_quantity=0
		astation=split(STATIONS_INDEX,",")
		set rsS=server.CreateObject("adodb.recordset")
		 'to count total job quantity of multi stations
		SQLS="select 1,JA.ACTION_VALUE from JOB_ACTIONS JA inner join JOB_STATIONS JS on JA.JOB_NUMBER=JS.JOB_NUMBER and JA.STATION_ID=JS.STATION_ID and JA.JOB_TYPE=JS.JOB_TYPE and JA.REPEATED_SEQUENCE=JS.REPEATED_SEQUENCE where JA.JOB_NUMBER='"&session("JOB_NUMBER")&"' and JA.SHEET_NUMBER='" &session("SHEET_NUMBER")&"' and JA.JOB_TYPE='"&session("JOB_TYPE")&"' and JA.ACTION_ID='AC00000259' and (JA.STATION_ID='"&current_station_id&"' and JA.REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"' or JS.STATUS=2)"
		rsS.open SQLS,conn,1,3
		if not rsS.eof then
			while not rsS.eof
				total_quantity=total_quantity+cint(rsS("ACTION_VALUE"))
			rsS.movenext
			wend
		end if
		rsS.close
		'if total job quantity is large or equal to initial quantity of job, job is closed
		if cint(total_quantity)>=cint(rs("JOB_START_QUANTITY"))*(ubound(astation)+1) then 
			job_is_closed=true
			rs("STATUS")=1
			rs("CLOSE_TIME")=request.Form("station_close_time")
		end if
		'to close any stations which have not closed.		
		SQLS="select STATUS,CLOSE_TIME from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='" &session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATUS=0" 
		rsS.open SQLS,conn,1,3
		if not rsS.eof then
			while not rsS.eof
			rsS("STATUS")=2
			rsS("CLOSE_TIME")=now()
			rsS.update
			rsS.movenext
			wend
		end if
		rsS.close
		set rsS=nothing
	end if
		
	rs("PREVIOUS_STATION_CLOSE_TIME")=request.Form("station_close_time")
	rs("PREVIOUS_STATION_ID")=current_station_id
	rs("PREVIOUS_STATION_TRANSACTION")=current_station_transaction
	if  isnull(rs("FINISHED_STATIONS_ID")) or rs("FINISHED_STATIONS_ID")="" then
		rs("FINISHED_STATIONS_ID")=rs("FINISHED_STATIONS_ID")&current_station_id&","
	else
		if instr(rs("FINISHED_STATIONS_ID"),current_station_id)=0 then
			rs("FINISHED_STATIONS_ID")=rs("FINISHED_STATIONS_ID")&current_station_id&","
		end if 
	end if
	
	total_defect_quantity = total_defect_quantity + clng(rs("JOB_DEFECTCODE_QUANTITY"))
	good_quantity = clng(rs("JOB_START_QUANTITY"))-total_defect_quantity
	rs("JOB_DEFECTCODE_QUANTITY")=total_defect_quantity
	rs("JOB_GOOD_QUANTITY")=good_quantity		
    rs("JOB_GOOD_QTY")=	good_quantity
	rs("JOB_ASSEMBLY_YIELD")=good_quantity/clng(rs("JOB_START_QUANTITY"))
	rs.update
end if
rs.close
'update master table
if job_is_closed=true then	
	'update by jack Zhang 2010-8-13 only for Normal Quantity and Rework Quantity
	SQL="select nvl(sum(JOB_GOOD_QUANTITY),0) as SUM_JOB_GOOD_QUANTITY,nvl(sum(JOB_DEFECTCODE_QUANTITY),0) as SUM_JOB_DEFECTCODE_QUANTITY,sum(decode(job_type,'R',job_start_quantity,0)) as RJobQty from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and Job_type<>'S'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		sum_job_good_quantity=csng(rs("SUM_JOB_GOOD_QUANTITY"))-csng(rs("RJobQty"))
		sum_job_defectcode_quantity=csng(rs("SUM_JOB_DEFECTCODE_QUANTITY"))
	end if
	rs.close
	SQL="select ASSEMBLY_GOOD_QUANTITY,DEFECTCODE_QUANTITY,ASSEMBLY_YIELD from JOB_MASTER where JOB_NUMBER='"&session("JOB_NUMBER")&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		rs("ASSEMBLY_GOOD_QUANTITY")=sum_job_good_quantity
		rs("DEFECTCODE_QUANTITY")=sum_job_defectcode_quantity
		rs("ASSEMBLY_YIELD")=sum_job_good_quantity/(sum_job_good_quantity+sum_job_defectcode_quantity)
		rs.update
	end if
	rs.close
	
	'disable tray id and job link
	SQL="update job_tray_mapping set is_disable =1 where job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' "
	rs.open SQL,conn,1,3
	
end if

'add by jack zhang 2011-1-17 do auto scrap
SQL="select * from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='" & trim(cstr(session("SHEET_NUMBER")))&"' and TRIM(JOB_TYPE)='"&trim(session("JOB_TYPE"))&"'"
set rsJobInfo=server.createobject("adodb.recordset")
rsJobInfo.open sql,conn,1,3
if rsJobInfo.recordcount>0 then	
	if cdbl(rsJobInfo("Job_Good_Quantity"))=0 then
		part_number_tag=rsJobInfo("PART_NUMBER_TAG")	
		IF (trim(session("JOB_TYPE"))="C") THEN
			SQL="SELECT * FROM CHANGE_MODEL_MAPPING WHERE NEW_MODEL='"+part_number_tag+"'"
			set rsChangeModelName=server.createobject("adodb.recordset")
			rsChangeModelName.open SQL,conn,1,3
			if rsChangeModelName.recordcount>0 then
				part_number_tag=rsChangeModelName("OLD_MODEL")
			end if
			rsChangeModelName.close 
		END IF 
		factory=rsJobInfo("FACTORY_ID")
		line_name=rsJobInfo("LINE_NAME")
		Qty=rsJobInfo("JOB_START_QUANTITY")
		jobType=rsJobInfo("JOB_TYPE")
		sheetNumber=rsJobInfo("SHEET_NUMBER")
		rsJobInfo.close
				 
		SQL="select * from JOB_MASTER_SCRAP_PRE WHERE REASON='"+trim(cstr(session("SHEET_NUMBER")))+"' and JOB_NUMBER='"+session("JOB_NUMBER")+"'"
		set rsReworkJobAutoScrap=server.createobject("adodb.recordset")
		rsReworkJobAutoScrap.open SQL,conn,1,3
		if rsReworkJobAutoScrap.recordcount=0 then		
			'save store info in job_master_store_pre
			set rsStore=server.createobject("adodb.recordset")
			SQL="SELECT * FROM JOB_MASTER_STORE_PRE where 1=2"
			rsStore.open SQL,conn,1,3
			rsStore.addnew
			storeId="WH"&NID_SEQ("PROD_STORE")
			rsStore("NID")=storeId
			rsStore("JOB_NUMBER")=session("JOB_NUMBER")
			rsStore("PART_NUMBER_TAG")=part_number_tag
			rsStore("FACTORY_ID")=factory_id
			rsStore("LINE_NAME")=line_name
			rsStore("STORE_CODE")=session("CODE")+ "-0000"
			rsStore("STORE_TIME")=now()
			rsStore("STORE_QUANTITY")=0
			rsStore("INSPECT_QUANTITY")=0
			rsStore("MRB")=0
			rsStore("MRB_RFP")=0
			rsStore("NO_YIELD")=0
			rsStore("YIELD")=0
			rsStore("SUB_JOB_NUMBERS")=session("JOB_NUMBER")&"-"&string(3-len(sheetNumber),"0")&sheetNumber
			rsStore("STORE_TYPE")=jobType
			if jobType="N" then
				rsStore("INPUT_QUANTITY")=Qty			
			end if
			rsStore("NOTE")="Main Line Auto Scrap"
			rsStore.update
			rsStore.close
		
			rsReworkJobAutoScrap.addnew
			NID="CH"&NID_SEQ("PROD_SCRAP")
			rsReworkJobAutoScrap("NID")=NID
			rsReworkJobAutoScrap("JOB_NUMBER")=session("JOB_NUMBER")
			rsReworkJobAutoScrap("PART_NUMBER_TAG")=part_number_tag
			rsReworkJobAutoScrap("FACTORY_ID")=factory
			rsReworkJobAutoScrap("LINE_NAME")=line_name
			rsReworkJobAutoScrap("SCRAP_CODE")=session("CODE") & "-0000"
			rsReworkJobAutoScrap("SCRAP_TIME")=now()
			rsReworkJobAutoScrap("SCRAP_QUANTITY")=Qty
			rsReworkJobAutoScrap("NOTE")="Main Line Auto Scrap"
			rsReworkJobAutoScrap("REASON")=trim(cstr(session("SHEET_NUMBER")))
			rsReworkJobAutoScrap("scrap_account")="142987"
			rsReworkJobAutoScrap("SCRAP_REASON_ID")="8"
			'Get part number and serial name for job
			SQLScrapRef="SELECT SERIES_NAME FROM SERIES_NEW WHERE NID=(SELECT SERIES_GROUP_ID FROM PRODUCT_MODEL WHERE ITEM_NAME=(SELECT PART_NUMBER_TAG FROM JOB_MASTER WHERE JOB_NUMBER='"&session("JOB_NUMBER")&"'))"
			set rsSerial=server.createobject("adodb.recordset")
			rsSerial.open SQLScrapRef,conn,1,3
			if rsSerial.recordcount=1 then
				Scrap_Ref="SP-A01-"&rsSerial("SERIES_NAME")&"-00"
			else
				Scrap_Ref="No Serial Name"
			end if
			rsSerial.close
			set rsSerial=nothing
			rsReworkJobAutoScrap("SCRAP_REFERENCE")=Scrap_Ref
			rsReworkJobAutoScrap("STORE_NID")=storeId
			rsReworkJobAutoScrap("PRINT_TIMES")=0
			rsReworkJobAutoScrap.update
			rsReworkJobAutoScrap.close
			
			SQL="update JOB_MASTER set FINAL_SCRAP_QUANTITY=FINAL_SCRAP_QUANTITY+"+cstr(Qty)+" where job_number='"+cstr(session("JOB_NUMBER"))+"'"
			set rsMaster=server.createobject("adodb.recordset")
			rsMaster.open SQL,conn,1,3
			
			SQL="update JOB_MASTER set STORE_STATUS='1' where job_number='"+cstr(session("JOB_NUMBER"))+"' "
			SQL=SQL+" and start_quantity = confirm_good_quantity+final_scrap_quantity"
			set rsMaster2=server.createobject("adodb.recordset")
			rsMaster2.open SQL,conn,1,3
			
			'add by jack zhang 2011-2-18
			SQL="update JOB set STATUS='1',close_time='"+cstr(now())+"' where job_number='"+cstr(session("JOB_NUMBER"))+"' "
			SQL=SQL+" and SHEET_NUMBER='" &trim(cstr(session("SHEET_NUMBER")))&"'"
			set rsJobInfo2=server.createobject("adodb.recordset")
			rsJobInfo2.open SQL,conn,1,3
			
			job_is_closed=true
		end if 
	end if 
end if 	
'end add
		
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" --> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=3;
adCount=0;
function showPopup()
{
	adCount+=1
	if(adCount<timePopup)
	{
		setTimeout("showPopup()",1000);
		document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
	else
	{
		closePopup()
	}
}
function closePopup()
{
	location.href="Station_Close.asp"
}
</script>
</head>

<body onLoad="showPopup();"  bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td width="100%" height="20"><div align="center" class="strongred">System will close window in 3 seconds.&nbsp;<br>
      系统将在3秒钟后关闭窗口。 <span id="countinsert"></span></div></td>
  </tr>
  <tr>
    <td height="20" class="t-t-DarkBlue">Job's notification 提示</td>
  </tr>
  <%if job_is_closed=true then%>
  <tr>
    <td height="20">Job is closed. 本工单已结束。 </td>
  </tr>
  <%else%>
  <tr>
    <td height="20">Next Station is 
    下一站是： <% =next_station_name%><%=next_station_chinese_name%></td>
  </tr>
  <%end if
  if new_rework=true then%>
  <tr>
    <td height="20">
    Rework工单是： <% =rework_job_number%></td>
  </tr>
  <%end if%>
</table>
</body>
</html>
<%session.Abandon()%>