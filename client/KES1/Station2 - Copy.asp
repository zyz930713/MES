<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Functions/PartCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/StationDisplay.asp" -->
<!--#include virtual="/Functions/GetJobStation.asp" -->
<!--#include virtual="/Functions/GetLeadTime.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->


<%
pagename="Station2.asp"
session("NEW_JOB")="false"
set rsJ=server.CreateObject("adodb.recordset")
set rsL=server.CreateObject("adodb.recordset")
set rsP=server.CreateObject("adodb.recordset")
set rsSH=server.CreateObject("adodb.recordset")

code=ucase(trim(request.Form("code")))
'get Job Number
inputJob=trim(request.Form("jobnumber"))
if instr(inputJob,"-")=0 then
	response.Redirect("Station1.asp?errorstring=Job Number is error, please scan sub job！<br>工单错误，请扫描子工单！")
end if
ajobnumber=split(inputJob,"-")
strSheetNo=ajobnumber(ubound(ajobnumber))

if len(strSheetNo)>3 then
	response.Redirect("Station1.asp?errorstring=Job Number is error, please scan sub job！<br>细分工单单编号超过3位错误，请扫描子工单！")
end if

if isnumeric(strSheetNo) then
	sheetnumber = cint(strSheetNo)
	if(sheetnumber<500) then
		jobtype="N"
	elseif(sheetnumber<600) then
		jobtype="R"
	elseif(sheetnumber<700) then
		jobtype="S"
	else
		jobtype="N"
	end if	
else
	response.Redirect("Station1.asp?errorstring=Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！")
end if
jobnumber=left(inputJob,len(inputJob)-4)

'check whether user is recorded operator.
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
	rs.open SQL,conn,1,3
	if rs.eof then
		codevalid=false
		valid=false
		word="Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 不存在，请联系工程师。"
	else
		if rs("LOCKED")="1" then
			codevalid=false
			valid=false
			word="Operator Code of "&code&" is locked, please contact engineer.<br>"&code&" 被锁定，请联系工程师。"
		else
			if rs("PRACTISED")="1" and (date()<rs("PRACTISE_START_TIME") or date()>rs("PRACTISE_END_TIME")) then
				codevalid=false
				valid=false
				word="Practicing period of Operator Code of "&code&" is expired, please contact engineer.<br>"&code&" 的实习期限已过期，请联系工程师。"
			else
			codevalid=true
			session("operator")=rs("OPERATOR_NAME")
			AUTHORIZED_STATIONS_ID=rs("AUTHORIZED_STATIONS_ID")
			AUTHORIZED_PARTS_ID=rs("AUTHORIZED_PARTS_ID")
			end if
		end if
	end if
rs.close

'job is shift in
if request.Form("shiftin")="1" then
	SQL="select J.STATUS,J.SHIFT_IN_PERSON,J.SHIFT_IN_TIME from JOB J where J.STATUS=5 and J.JOB_NUMBER='"&jobnumber&"' and J.SHEET_NUMBER='"&sheetnumber&"' and J.JOB_TYPE='"&jobtype&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		rs("STATUS")=0
		rs("SHIFT_IN_PERSON")=rs("SHIFT_IN_PERSON")&code&","
		rs("SHIFT_IN_TIME")=rs("SHIFT_IN_TIME")&now&","
		rs.update
	else
		codevalid=false
		valid=false
		word="Job of "&inputJob&" is not shift-out, can not be shift-in!<br>"&inputJob&"工单没有被停线，不能开线！"
	end if
	rs.close
end if

if codevalid=true then
	'get job's property
	SQL="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.STATUS,J.PART_NUMBER_ID,J.PART_NUMBER_TAG,J.FACTORY_ID,J.LINE_NAME,J.STATIONS_ROUTINE,"
	SQL=SQL+ " J.JOB_START_QUANTITY,J.START_TIME,J.FIRST_STATION_ID,J.CURRENT_STATION_ID,J.LAST_STATION_ID,J.FINISHED_STATIONS_ID,J.STATIONS_INDEX,OPENED_STATIONS_ID, "
	SQL=SQL+ "SUBSTR(STATIONS_INDEX,LENGTH(OPENED_STATIONS_ID)+1,10) AS OPEN_STATION,J.STATIONS_TRANSACTION "
	SQL=SQL+ "from JOB J where J.JOB_NUMBER='"&jobnumber&"' and J.SHEET_NUMBER='"&sheetnumber&"' and J.JOB_TYPE='"&jobtype&"'"

	rs.open SQL,conn,1,3
	if rs.eof then'no such job, it is a new job.
		station_status=0'station is new opened
		if jobtype="R" or left(inputJob,3)="RWK" or left(inputJob,3)="TKB" then 'retest job
			SQL="select a.quantity,b.station_name from print_newjob_history a,rework_printing b where a.new_jobnumber = get_sub_job_number(b.job_number,b.seq) and a.new_jobnumber = '"&inputJob&"' "
			rsSH.open SQL,conn,1,3
			if not rsSH.eof then
				current_station_id=rsSH("station_name")
				job_start_quantity=rsSH("quantity")	
				'get information from job_master
				SQL="select * from job_master where job_number = '"&jobnumber&"' "
				rsJ.open SQL,conn,1,3
				if not rsJ.eof then
					part_number_id=rsJ("PART_NUMBER_ID")
					SQLP="select STATIONS_INDEX,STATIONS_TRANSACTION,STATIONS_ROUTINE from PART where NID='"&part_number_id&"'"
					rsP.open SQLP,conn,1,3
					if not rsP.eof then
						stations_index=rsP("STATIONS_INDEX")
						stations_transaction=rsP("STATIONS_TRANSACTION")
						stations_routine=rsP("STATIONS_ROUTINE")
					end if
					rsP.close					
					stations_index=right(stations_index,len(stations_index)-instr(stations_index,current_station_id)+1)

					'check current station
					word = checkCurrentStaion(current_station_id,"",Request.Cookies("current_station_id"))
					if len(word)>10 then					
						valid=false
					else
						word = ""
						station_status=0
						factory_id=rsJ("FACTORY_ID")						
						part_number_tag=rsJ("PART_NUMBER_TAG")
						line_name=rsJ("LINE_NAME")		
						opened_stations_id=""
						finished_stations_id=""						
						astation=split(stations_index,",")
						atransaction=split(stations_transaction,",")
						stationsTransaction=""
						for i=0 to ubound(astation)
							stationsTransaction=atransaction(ubound(atransaction)-i)&","&stationsTransaction
						next

						if stations_routine="0" and ubound(astation)>1 then'fixed routine
							display_stations_id=current_station_id&","&astation(1)
						else
							display_stations_id=stations_index
						end if
										
						session("JOB_NUMBER")=jobnumber
						session("SHEET_NUMBER")=sheetnumber
						session("JOB_TYPE")=jobtype						
						session("PART_NUMBER_ID")=part_number_id
						session("PART_NUMBER_TAG")=part_number_tag
						session("FACTORY_ID")=factory_id
						session("LINE_NAME")=line_name
						session("STATIONS_ROUTINE")=stations_routine
						session("JOB_START_QUANTITY")=job_start_quantity
						session("FIRST_STATION_ID")=current_station_id						
						session("CURRENT_STATION_ID")=current_station_id
						session("LAST_STATION_ID")=astation(ubound(astation))
						session("STATIONS_INDEX")=stations_index
						session("STATIONS_TRANSACTION")=left(stationsTransaction,len(stationsTransaction)-1)
						session("LEAD_TIME")=getLeadTime(jobnumber,part_number_tag,factory_id)
						session("JOB_PRIORITY")=5
						session("NEW_JOB_TYPE")="REWORK"
						session("NEW_JOB")="true"
						session("CODE")=code
						valid=true
					end if
				end if
				rsJ.close
			else
				valid=false
				word="Info about Job of "&inputJob&" has error, please contact engineer!<br>"&inputJob&"工单信息错误<子工单没有打印>，请联系生产工程师！"
			end if
			
			rsSH.close
			
		elseif jobtype="N" or jobtype="S" or jobtype="C" then 'job is normal job
			'get Part info and line info from Prod
			SQLSH="select SheetType from tbl_MES_LotMasterSub where WipEntityName='"&jobnumber&"' and SheetNumber="
			if(jobtype="N") then
				SQLSH=SQLSH&sheetnumber
			else
				SQLSH=SQLSH&"1"
			end if		
			rsSH.open SQLSH,connTicket,1,3
		
			if not rsSH.eof then'
				sheettype=cint(rsSH("SheetType"))				
				SQLJ="select M.OrganizationId,M.PrimaryItemDesc,M.LineCode,M.StartQuantity,M.DateCreation,M.CreatePerson,M.DateLastUpdate,M.LastUpdatePerson,M.BOM_LABEL,"
				SQLJ=SQLJ+ " U.CODE as CREATE_CODE,U.NAME as CREATE_NAME,JOB_PRIORITY from tbl_MES_LotMaster M left join tbl_User_ID U on M.CREATEPERSON=U.ID where WipEntityName='"&jobnumber&"'"
				rsJ.open SQLJ,connTicket,1,3
				
				if not rsJ.eof then
					organizationid=rsJ("OrganizationId")
					partnumber=rsJ("PrimaryItemDesc")
					bomlabel=rsJ("BOM_LABEL")
				
					'ADD FOR CHANGE MODEL
					IF (jobtype="C") THEN
						SQLMMM="SELECT NEW_MODEL FROM CHANGE_MODEL_MAPPING WHERE OLD_MODEL='"+partnumber+"'"
						set rsCM=server.CreateObject("adodb.recordset")
						rsCM.open SQLMMM,conn,1,3
						if (rsCM.recordcount>0) then
							partnumber=rsCM("NEW_MODEL")
						end if
						rsCM.close 
					END IF 
					linename=rsJ("LineCode")
					totalstartquantity=rsJ("StartQuantity")
					erp_create_time=rsJ("DateCreation")
					erp_create_person=rsJ("CreatePerson")
					erp_last_update_time=rsJ("DateLastUpdate")
					erp_last_update_person=rsJ("LastUpdatePerson")
					createcode=rsJ("CREATE_CODE")
					createname=rsJ("CREATE_NAME")
					JOB_PRIORITY=rsJ("JOB_PRIORITY")
				
					SQLL="select NID from LINE where STATUS=1 and LINE_NAME='"&linename&"'"
					rsL.open SQLL,conn,1,3
					if rsL.eof then'Line name inputted is not existed.
						valid=false
						word="Line Name of "&linename&" is not existed, please contact engineer.<br>"&linename&" 线不存在，请联系工程师"
					else
						line_id=rsL("NID")
						boolean_partcheck=false
						SQLP="select a.NID,a.PART_RULE,a.FACTORY_ID,a.LINES_INDEX,a.STATIONS_INDEX,a.STATIONS_TRANSACTION,a.STATIONS_ROUTINE,a.INITIAL_QUANTITY "
						SQLP=SQLP+ "from PART a , routing b"
						SQLP=SQLP+" where a.mother_routing_id=b.nid and a.STATUS=1 and a.ROUTINE_TYPE=0 and a.PART_RULE is not null and mother_routing_id is not null "
						SQLP=SQLP+" and length(mother_routing_id)>0 "
					'	SQLP=SQLP+" and decode(b.part_type,'0','N','1','R','5','S','4','C')='"+jobtype+"' and b.IS_DELETE='0' order by a.MEET_PRIORITY, a.PART_RULE desc"
			SQLP=SQLP+" and decode(b.part_type,'0','N','1','R','5','S','4','C')='"+jobtype+"' and b.IS_DELETE='0' and A.PART_RULE like '%"&partnumber&"%' and a.LINES_INDEX like '%"&line_id&"%' order by a.MEET_PRIORITY, a.PART_RULE desc"
					rsP.open SQLP,conn,1,3'part number inputted is correct.
						do while not rsP.eof
							if not isnull(rsP("PART_RULE")) and rsP("PART_RULE")<>"" then
								a_part_rule=split(rsP("PART_RULE"),",")' sparate PART_RULE to single rule
								for p=0 to ubound(a_part_rule)'loop single rule
									boolean_partcheck=PartCheck(trim(a_part_rule(p)),partnumber)
									if boolean_partcheck=true then
										if isnull(rsP("LINES_INDEX")) or instr(rsP("LINES_INDEX"),line_id)<>0 then' check does line name is allowed to produce						
											if instr(AUTHORIZED_PARTS_ID,rsP("NID"))<>0 then
												valid=true
												part_number_id=rsP("NID")
												part_number_tag=partnumber
												factory_id=rsP("FACTORY_ID")
			
												stations_routine=rsP("STATIONS_ROUTINE")
												job_start_quantity=rsP("INITIAL_QUANTITY")
												stations_index=rsP("STATIONS_INDEX")
												stations_transaction=rsP("STATIONS_TRANSACTION")
												astation=split(stations_index,",")
												atransaction=split(stations_transaction,",")
												current_station_id=astation(0)
												'check current station
												word = checkCurrentStaion(current_station_id,"",Request.Cookies("current_station_id"))
												if len(word)>10 then
													valid=false
												else
												
													display_stations_id=current_station_id												
													finished_stations_id=""
													display_stations_id=display_stations(astation,atransaction,display_stations_id)
													
													if sheettype=0 then'if sheet is not conjuctive
														last_station_id=last_for_notconj(astation,atransaction)
													else'if sheet is conjuctive
														last_station_id=last_for_conj(astation,atransaction)
														stations_refresh astation,atransaction,last_station_id,stations_index,stations_transaction
													end if
													
													session("NEW_JOB")="true"
													session("CODE")=code
													session("ORGANIZATION_ID")=organizationid
													session("JOB_NUMBER")=jobnumber
													session("SHEET_NUMBER")=sheetnumber
													session("JOB_TYPE")=jobtype
													session("PART_NUMBER_ID")=part_number_id
													session("PART_NUMBER_TAG")=part_number_tag
													session("LEAD_TIME")=getLeadTime(jobnumber,part_number_tag,factory_id)
													session("FACTORY_ID")=factory_id
													session("LINE_NAME")=linename
													session("STATIONS_ROUTINE")=stations_routine
													session("JOB_TOTAL_START_QUANTITY")=totalstartquantity
												
													session("CREATE_CODE")=createcode
													session("CREATE_NAME")=createname
													session("JOB_START_QUANTITY")=job_start_quantity
	
													session("FIRST_STATION_ID")=astation(0)
													session("CURRENT_STATION_ID")=astation(0)
													session("LAST_STATION_ID")=last_station_id
						
													session("STATIONS_INDEX")=stations_index
													session("STATIONS_TRANSACTION")=stations_transaction
													session("ERP_CREATE_TIME")=erp_create_time
													session("ERP_CREATE_BY")=erp_create_person
													session("ERP_LAST_UPDATE_TIME")=erp_last_update_time
													session("ERP_LAST_UPDATE_BY")=erp_last_update_person
													session("JOB_PRIORITY")=JOB_PRIORITY
													session("BOM_LABEL")=bomlabel
													
													IF (jobtype="N") THEN
														if instr(jobnumber,"E")<>0 then
															session("NEW_JOB_TYPE")="ENG"
														else
															session("NEW_JOB_TYPE")="NORMAL"
														end if 
													END IF 
													IF (jobtype="R") THEN
															session("NEW_JOB_TYPE")="REWORK"
													END IF 
													IF (jobtype="C") THEN
															session("NEW_JOB_TYPE")="CHANGE MODEL"
													END IF 
													
													if stations_routine="1" then
														display_stations_id=stations_index
													end if
												end if
											else
												word="You do not have right to do products of Part Number of "&partnumber&", "
												word=word+" please contact engineer.<br>你无权操作"&partnumber&"，请联系工程师。"
											end if
										else
											valid=false
											word=linename&" line cannnot operate this Part "&a_part_rule(p)&", "
											word=word+" please contact engineer.<br>"&linename&"线别无权操作此型号，请联系工程师。"
										end if
									exit for
									end if
								next
								if boolean_partcheck=true then
									exit do
								end if
							end if
						rsP.movenext
						loop
						rsP.close
						if boolean_partcheck=false then
							valid=false
							word="Part Number of "&partnumber&" is not existed, please contact engineer.<br>"&partnumber&" 不存在，请联系工程师。"
						end if
					end if
					rsL.close
				else
					valid=false
					word="Info about Job of "&inputJob&" has error, please contact engineer!<br>"&inputJob&"工单信息错误<Prod没有相应工单>，请联系生产工程师！[X00001]"
				end if
				rsJ.close
			else
				valid=false
				word="Info about Job of "&inputJob&" has error, please contact engineer!<br>"&inputJob&"工单信息错误<子工单没有打印>，请联系生产工程师！[X00002]"
			end if
			rsSH.close						
		end if
	else'it is opened job.
		session("NEW_JOB")="false"
		if rs("STATUS")="0" then ' job is opened	
			valid=true
			sheetnumber=rs("SHEET_NUMBER")
			jobtype=rs("JOB_TYPE")
			factory_id=rs("FACTORY_ID")
			start_time=rs("START_TIME")
			current_station_id=rs("CURRENT_STATION_ID")
			part_number_id=rs("PART_NUMBER_ID")
			part_number_tag=rs("PART_NUMBER_TAG")
			line_name=rs("LINE_NAME")
			stations_routine=rs("STATIONS_ROUTINE")
			job_start_quantity=rs("JOB_START_QUANTITY")	
			opened_stations_id=rs("OPENED_STATIONS_ID")	
			finished_stations_id=rs("FINISHED_STATIONS_ID")
			if isnull(rs("STATIONS_INDEX")) or rs("STATIONS_INDEX")="" then
				SQLP="select STATIONS_INDEX,STATIONS_TRANSACTION,STATIONS_ROUTINE from PART where NID='"&part_number_id&"'"
				rsP.open SQLP,conn,1,3
				if not rsP.eof then
					stations_index=rsP("STATIONS_INDEX")
					stations_transaction=rsP("STATIONS_TRANSACTION")
				end if
				rsP.close
			else
				stations_index=rs("STATIONS_INDEX")
				stations_transaction=rs("STATIONS_TRANSACTION")
			end if
			
			'check current station
			word = checkCurrentStaion(rs("OPEN_STATION"),rs("CURRENT_STATION_ID"),Request.Cookies("current_station_id"))
			if len(word)>10 then
				valid=false
			else
				current_station_id = word
				word = ""
				session("CURRENT_STATION_ID")=current_station_id
				session("CODE")=code
				session("JOB_NUMBER")=jobnumber
				session("SHEET_NUMBER")=sheetnumber
				session("JOB_TYPE")=jobtype
				session("FACTORY_ID")=factory_id
				session("PART_NUMBER_ID")=part_number_id
				session("PART_NUMBER_TAG")=part_number_tag
				session("LINE_NAME")=line_name
				session("STATIONS_ROUTINE")=stations_routine
				
				astation=split(stations_index,",")
				atransaction=split(stations_transaction,",")
				if stations_routine="0" then'fixed routine
					befStationIndex=0
					aftStationIndex=0
					for i=0 to ubound(astation)'get three stations to display
						if i>0 then						
							if atransaction(i-1)="0" or atransaction(i-1)="2" then							
								befStationIndex=i-1
							end if								
						end if
						if i<ubound(astation) then
							if atransaction(i+1)="0" or atransaction(i+1)="2" then							
								aftStationIndex=i+1
							end if
						end if
						if current_station_id=astation(i) then							
							display_stations_id=current_station_id
							if befStationIndex < i then
								display_stations_id = astation(befStationIndex)&","&display_stations_id
							end if
							if aftStationIndex > i then
								display_stations_id = display_stations_id &","& astation(aftStationIndex)
							end if
							exit for							
						end if
					next
				else'
					display_stations_id=stations_index
				end if
	
				SQLP="select STATUS from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' "
				SQLP=SQLP+" and STATION_ID='"&current_station_id&"'"
				rsP.open SQLP,conn,1,3
				if rsP.eof then 'no record, this station is new opened
					station_status=0
				else
					station_status=1
				end if
				rsP.close
			end if
		elseif rs("STATUS")="1" then'job is closed.
			valid=false
			word="Job of "&inputJob&" is closed!<br>"&inputJob&"工单已结束！"
			para=""
		elseif rs("STATUS")="2" then'job is paused.
			valid=false
			word="Job of "&inputJob&" is paused, please contact engineer!<br>"&inputJob&"工单已被暂停，请联系生产工程师！"
			para=""
		elseif rs("STATUS")="3" then'job is locked.
			valid=false
			word="Job of "&inputJob&" is locked, please contact engineer!<br>"&inputJob&"工单已被锁定，请联系生产工程师！"
			para=""
		elseif rs("STATUS")="4" then'job is aborted.
			valid=false
			word="Job of "&inputJob&" is aborted, please contact engineer!<br>"&inputJob&"工单已被废除，请联系生产工程师！"
			para=""
		elseif rs("STATUS")="5" then'shift is stoped.
			valid=false
			word="Job of "&inputJob&" is shift-out, please contact engineer!<br>"&inputJob&"工单已被停线，需要领班开线后继续！"
			para="&shift=Y"
		end if
	end if
	rs.close
end if
%>

<%	
function checkCurrentStaion(canOpenId,canCloseId,cookiesStationId)
	session("checkCurrentStaion")=""
	SQL = "select nid from station where mother_station_id = '"&cookiesStationId&"' and nid in ('"&replace(stations_index,",","','")&"') "	
	set rsChkSt = server.createobject("adodb.recordset")
	rsChkSt.open SQL,conn,1,3
	if not rsChkSt.eof then
		if instr(finished_stations_id,rsChkSt("nid")) > 0 then
			checkCurrentStaion = "The current station is closed in the routing of this job number.<br>该工单的制程已经完成当前站点。"
		elseif instr(opened_stations_id,rsChkSt("nid"))=0 and canOpenId <> rsChkSt("nid") then
			checkCurrentStaion = "Can not open this station, because previous station is not opened.<br>前一站未开，不能开当前站点。"
		'elseif opened_stations_id<>"" and  instr(opened_stations_id,rsChkSt("nid")) = 0 and canCloseId <> rsChkSt("nid") then
		'    checkCurrentStaion = "Can not close this station, because previous station is not Open.<br>前一站未关，不能开当前站点。"
		elseif instr(opened_stations_id,rsChkSt("nid")) > 0 and canCloseId <> rsChkSt("nid") then
			'checkCurrentStaion = "Can not close this station, because previous station is not closed.<br>前一站未关，不能关当前站点。"
			checkCurrentStaion = rsChkSt("nid")
			session("checkCurrentStaion")="Can not close this station, because previous station is not closed.<br>前一站未关，不能关当前站点。"
		else
			checkCurrentStaion = rsChkSt("nid")
		end if
	else
		checkCurrentStaion = "The routing of this job number does not contain your current station.<br>该工单的制程不包含您当前的工作站。"
	end if
	rsChkSt.close
	set rsChkSt = nothing
end function 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KES1 Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<%
if valid=false then
	response.Redirect("Station1.asp?errorstring="&word&para)
else
%>
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="javascript">
function pageload()
{
	with(document.form1)
	{
		if(Next.disabled==false)
		{
			Next.focus();
		}
 	}
}

function submitonce(theform)
{
	if (document.all||document.getElementById)
	{
		for (i=0;i<document.form1.length;i++)
		{
			var tempobj=document.form1.elements[i]
			if(tempobj.type.toLowerCase() =="submit"||tempobj.type.toLowerCase()=="reset")
			tempobj.disabled=true
		}
	}
}
</script>
</head>
<body onLoad="pageload();loadimage()" onKeyPress="keyhandler()"  bgcolor="#339966">
<form action="/KES1/Station3.asp" method="post" name="form1" target="_self" onsubmit="return submitonce()">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="4" class="t-t-DarkBlue">Confirm the station
      确认工作站</td>
  </tr>
  <tr>
    <td height="20" colspan="4" class="t-t-Borrow">Operator 操作员:  
      <% =session("operator") %> (<% =session("CODE") %>)</td>
  </tr>
  <tr>
    <td width="32%">Job Number 工单号</td>
    <td width="22%"><%=inputJob%></td>
    <td width="26%" height="20">Part Number 型号 </td>
    <td width="20%" height="20"><% =session("PART_NUMBER_TAG") %></td>
    </tr>
  <tr>
    <td>Line Name 线别 </td>
    <td><% =session("LINE_NAME") %></td>
    <td height="20">Job Start Time 工单开始时间 </td>
    <td height="20"><% =start_time %></td>
  </tr>
  </table>
  <table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-t-Borrow">Stations Status </td>
    </tr>
  <tr class="t-t-FIN">
    <td height="20"><div align="center">NO. 序号</div></td>
    <td height="20"><div align="center">Name 站名</div></td>
    <td height="20"><div align="center">Status 状态</div></td>
    <td height="20"><div align="center">Operator操作员</div></td>
    <td><div align="center">Elapsed Time 用时</div></td>
  </tr>
  <%
  SQL="select NID,STATION_NAME,STATION_CHINESE_NAME,MOTHER_STATION_ID from STATION where NID in ('"&replace(display_stations_id,",","','")&"')"

  rs.open SQL,conn,1,3
  if not rs.eof then
  i=1
  for j=0 to ubound(astation)
	  while not rs.eof
		operatorcode=null
		station_start_time=null
		station_close_time=null
		unit=null
		stationStatus=0
		if rs("NID")=astation(j)then
			if stations_routine="0" then
			    classtype=""
				if instr(opened_stations_id,rs("NID"))<>0 then
					getJobStation jobnumber,sheetnumber,jobtype,rs("NID"),operatorcode,station_start_time,station_close_time					
					unit="m"
					stationStatus=1
				end if				
				if instr(finished_stations_id,rs("NID"))<>0 then
					classtype="StrongGreen"
					stationStatus=2
				elseif rs("NID")=current_station_id then
					classtype="strongred"
					if station_status=0 then
						stationStatus=1
					 else
					 	stationStatus=2
					 end if
				else
					classtype=""
				end if
				
				if rs("NID")=current_station_id then
					if instr(AUTHORIZED_STATIONS_ID,rs("MOTHER_STATION_ID"))=0 or isnull(AUTHORIZED_STATIONS_ID) then
						stationright=false
						session("stationright")=false
						classtype="t-b-blue"
					else
						stationright=true
						session("stationright")=true
						session("CURRENT_STATION_NAME")=rs("STATION_NAME")&"&nbsp;"&rs("STATION_CHINESE_NAME")
					end if
				end if				
  %>
			   <tr class="<%=classtype%>">
				<td height="20"><div align="center"><%=i%></div></td>
				<td height="20"><%=rs("STATION_NAME")%>&nbsp;<%=rs("STATION_CHINESE_NAME")%>&nbsp;</td>
				<td height="20"><div align="center">
					  <%if stationStatus=1 then%>
					  Before<br>开始
					  <%elseif stationStatus=2 then%>
					  After<br>结束
					  <%else%>
					  &nbsp;
					  <%end if%>
				  </div></td>
				<td height="20"><div align="center"><%if stations_routine="0" then%><%=operatorcode%><%end if%>&nbsp;</div></td>
				<td><div align="center"> <%=datediff("n",station_start_time,station_close_time)%><%=unit%>&nbsp;</div></td>
			  </tr>
  <%		else
				session("stationright")=true
				total_quantity=0
				SQLP="select ACTION_VALUE from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' and STATION_ID='"&rs("NID")&"' and ACTION_ID='AC00000259'" 'to count total job quantity of all stations
				rsP.open SQLP,conn,1,3
				if not rsP.eof then
					while not rsP.eof
						total_quantity=total_quantity+cint(rsP("ACTION_VALUE"))
						rsP.movenext
					wend
				end if
				rsP.close
				'if total job quantity in all stations is less than initial quantity of job and opertator has authority, new station can be create
				if cint(total_quantity)<cint(job_start_quantity) and instr(AUTHORIZED_STATIONS_ID,rs("NID"))<>0 and not isnull(AUTHORIZED_STATIONS_ID) then 
	%>
					<tr>
						<td height="20"><div align="center"><%=i%></div></td>
						<td height="20">
						<%if stations_routine="1" then%>
							<input name="sequence_type<%=i%>" type="hidden" value="new"><input name="station_id<%=i%>" type="hidden" value="<%=rs("NID")%>">
							<img src="" name="img<%=i%>" id="idimg<%=i%>" width="25" height="25" onClick="changeimage(<%=i%>)">New 
						<%end if%>
						<%=rs("STATION_NAME")%>&nbsp;<%=rs("STATION_CHINESE_NAME")%>&nbsp;</td>
						<td height="20"><div align="center">&nbsp;</div></td>
						<td height="20"><div align="center">&nbsp;</div></td>
						<td><div align="center">&nbsp;</div></td>
					</tr>
	  <%
				end if
			end if
		end if
	  rs.movenext
	  wend
  rs.movefirst
  i=i+1
  next
  end if
  rs.close  
  
  if stations_routine="1" then
  	SQL="select 1,JS.STATION_ID,JS.REPEATED_SEQUENCE,JS.OPERATOR_CODE,S.STATION_NAME,JA.ACTION_VALUE from JOB_STATIONS JS inner join STATION S on JS.STATION_ID=S.NID left join JOB_ACTIONS JA on (JS.JOB_NUMBER=JA.JOB_NUMBER and JS.SHEET_NUMBER=JA.SHEET_NUMBER and JS.JOB_TYPE=JA.JOB_TYPE and JS.STATION_ID=JA.STATION_ID and JS.REPEATED_SEQUENCE=JA.REPEATED_SEQUENCE) where JS.JOB_NUMBER='"&jobnumber&"' and JS.SHEET_NUMBER='"&sheetnumber&"' and JS.JOB_TYPE='"&jobtype&"' and JS.STATUS=1 and JA.ACTION_ID in ('AC00000241','AC00000240','AC00000242')"
	rs.open SQL,conn,1,3
	if not rs.eof then
  %>
	<tr>
	  <td height="5" colspan="5" class="today"></td>
    </tr>
	<%	while not rs.eof%>
	<tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><input name="sequence_type<%=i%>" type="hidden" value="old"><input name="station_id<%=i%>" type="hidden" value="<%=rs("STATION_ID")%>"><input name="repeated_sequence<%=i%>" type="hidden" value="<%=rs("REPEATED_SEQUENCE")%>"><img src="" name="img<%=i%>" id="idimg<%=i%>" width="25" height="25" onClick="changeimage(<%=i%>)"><%=rs("STATION_NAME")%>&nbsp;<%=rs("REPEATED_SEQUENCE")%>(<%=rs("ACTION_VALUE")%>)</td>
    <td height="20">After</td>
    <td height="20"><%=rs("OPERATOR_CODE")%></td>
    <td>&nbsp;</td>
  </tr>
  <%	rs.movenext
  		i=i+1
  	wend
  	end if
	rs.close
  end if%>
  <tr>
    <td height="20" colspan="5"><%if stationright=false and stations_routine="0" then%>
      <span class="strongred">You are not certificated operator<br>
      你无权操作此站</span>
      <%end if%>&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="5"><div align="center">
      <input name="selected_station" type="hidden" value="">
      <input name="Next" type="submit" <%if stationright=false and stations_routine="0" then%>disabled="true"<%end if%> value="Next 下一步">
&nbsp;
<input type="button" name="Button" value="Back 返回" onClick="javascript:location.href='/KES1/Station1.asp'">
&nbsp;
<%if (TRIM(session("JOB_TYPE"))="R" or TRIM(session("JOB_TYPE"))="C") then
	SQL="SELECT JOB_GOOD_QUANTITY FROM JOB WHERE JOB_NUMBER='"&jobnumber&"' AND SHEET_NUMBER='" &sheetnumber &"' AND FINISHED_STATIONS_ID IS NOT NULL"
	set rsJobInfo=server.createobject("adodb.recordset")	
	rsJobInfo.open SQL,conn,1,3
	ISSHOW="FALSE"
	if(rsJobInfo.recordcount>0) then
		if (rsJobInfo("JOB_GOOD_QUANTITY")="0") then
			ISSHOW="TRUE"
		end if 
	end if
	rsJobInfo.close
 	if ISSHOW="TRUE" then
%>
	<input type="button" name="Button" value="Close Job 关闭工单" onClick="javascript:location.href='/KES1/CloseReworkJob.asp'">
<%	end if
  end if
%>
    </div></td>
    </tr>
</table>
</form>
</body>
<%end if%>
</html>
<%set rs=nothing
set rsL=nothing
set rsP=nothing
set rsSH=nothing%>
<!--#include virtual="/Functions/JImage.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->