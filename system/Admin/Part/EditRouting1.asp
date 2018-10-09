<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
upcount=0
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
partnumber=trim(request.Form("partnumber"))
part_rule=trim(request.Form("part_rule"))
Guid=trim(request.Form("txtGuid"))
yield=request.Form("yield")
ROUTINGID=id
if yield="" or isnull(yield) or isempty(yield)  then
	yield=0
end if

SQL="select * from Routing where NID='"&id&"' and IS_DELETE<>1"
rs.open SQL,conn,1,3
if  rs.eof then
	word="Part of "&partnumber&" has not existed, please input again."
	action="history.back()"	
else
	rs("PART_NUMBER")=partnumber
	rs("JOB_RULE")=trim(request.Form("job_rule"))
	if right(part_rule,1)="," then
		this_part_rule=left(part_rule,len(part_rule)-1)
	else
		this_part_rule=part_rule
	end if
	rs("PART_RULE")=this_part_rule
	rs("PART_TYPE")=request.Form("partType")
	rs("FACTORY_ID")=request.Form("factory")
	rs("SECTION_ID")=request.Form("section")
	rs("STATIONS_COUNT")=request.Form("stationscount")
	lines_index=replace(request.Form("toitem1")," ","")
	rs("LINES_INDEX")=lines_index
	stations_index=replace(request.Form("toitem2")," ","")
	'old_stations_index=rs("STATIONS_INDEX")
	rs("STATIONS_INDEX")=stations_index
	rs("STATIONS_ROUTINE")=request.Form("stationroutine")
	rs("STATIONS_TRANSACTION")=getStationsTransaction_New(stations_index)
	rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
	if trim(request.Form("initial_quantity"))<>"" then
		rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
	else
		rs("INITIAL_QUANTITY")=0 
	end if
	rs("TARGET_YIELD")=yield
	rs("MEET_PRIORITY")=request.Form("priority")
	rs("IS_DELETE")=0
	rs("LASTUPDATE_PERSON")=session("code")
	rs("LASTUPDATE_TIME")=date()
	rs.update

	'Check Action and DefectCode
	stationString=Split(stations_index,",")
	set rsc=server.createobject("adodb.recordset")
	set rsd=server.createobject("adodb.recordset")
	
	'Delete old data in action
	set rsD=server.createobject("adodb.recordset")
	SQL1="delete from ROUTING_ACTION_DETAIL where ROUTING_ID='"&id&"'"
	rsD.open SQL1,conn,1,3	 
	SQL1="delete from ROUTING_DEFECT_DETAIL where ROUTING_ID='"&id&"'"
	rsD.open SQL1,conn,1,3
	For   i   =   Lbound(stationString)   to   Ubound(stationString) 			 
		SQL_CheckA="select * from ROUTING_ACTION_DETAIL_Temp where GUID='"&Guid&"' and STATION_ID='"&stationString(i)&"'"
		rsc.open SQL_CheckA,conn,1,3
		if rsc.recordcount=0 then
			word="Please set Stations' Action and DefectCode."
			action="location.href='/Admin/Part/EditRouting.asp?id="&id&"'" 
		else
			set rsA=server.createobject("adodb.recordset")
			SQL_CA="insert into ROUTING_ACTION_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE) SELECT 'RA'|| lpad(ROUTING_ACTION_SEQ.nextval,8,'0'), '"&id&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE FROM ROUTING_ACTION_DETAIL_TEMP where GUID='"&Guid&"' and station_id in ('"&stationString(i)&"')"
			rsA.open SQL_CA,conn,1,3
		end if
		rsc.close
		
		set rsP=server.createobject("adodb.recordset")
		SQL_CD="insert into ROUTING_DEFECT_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE) SELECT 'RD' || lpad(ROUTING_DEFECT_SEQ.nextval,8,'0'),'"&id&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE FROM ROUTING_DEFECT_DETAIL_TEMP where GUID='"&Guid&"' and  station_id in('"&stationString(i)&"')" 
		rsP.open SQL_CD,conn,1,3
		'end if
		'rsd.close
	next
	set rsc=nothing
	set rsd=nothing
	set rsP=nothing
	'delete the old data 
	set rs1=server.createobject("adodb.recordset")
	SQL1="delete from ROUTING_ACTION_DETAIL_Temp where guid='"&Guid&"'"
	rs1.open SQL1,conn,1,3
	
	SQL1="delete from ROUTING_DEFECT_DETAIL_Temp where guid='"&Guid&"'"
	rs1.open SQL1,conn,1,3
	
	if word="" and action="" then
		'Copy the data from new table to old table
		'get Part_ID
	    set rsPart=server.createobject("adodb.recordset")
		SQL="SELECT NID FROM PART WHERE Mother_Routing_ID='"+ROUTINGID+"'"
		
		rsPart.open SQL,conn,1,3
		if(rsPart.recordcount<>0) then
			Part_ID=rsPart("NID")
		else
				set rsRouting=server.createobject("adodb.recordset")
				Part_ID ="PA"&NID_SEQ("PART")
				SQL="INSERT INTO PART"
				SQL=SQL+"("
				SQL=SQL+" NID,PART_NUMBER,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID,INITIAL_QUANTITY,STATIONS_COUNT,MAX_INTERVAL,"
				SQL=SQL+" STATIONS_TRANSACTION,STATIONS_ROUTINE,LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,Mother_Routing_ID"
				SQL=SQL+" )"
				SQL=SQL+" select '"+Part_ID+"',PART_NUMBER,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID, "
				SQL=SQL+" INITIAL_QUANTITY,STATIONS_COUNT,MAX_INTERVAL,STATIONS_TRANSACTION,STATIONS_ROUTINE,"
				SQL=SQL+" LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,NID from ROUTING WHERE NID='"+ROUTINGID+"'"
				rsRouting.open SQL,conn,1,3
		end if
		set rsPart=nothing
		'1) UPDATE PART
		set rsRouting=server.createobject("adodb.recordset")
		SQL="update part "
		SQL=SQL+" set (PART_NUMBER,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID,INITIAL_QUANTITY,STATIONS_COUNT,STATIONS_INDEX,MAX_INTERVAL,STATIONS_TRANSACTION,"
		SQL=SQL+" STATIONS_ROUTINE,LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,Mother_Routing_ID)="
		SQL=SQL+" (select PART_NUMBER,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID,INITIAL_QUANTITY,STATIONS_COUNT,'',MAX_INTERVAL,STATIONS_TRANSACTION, "
		SQL=SQL+" STATIONS_ROUTINE,LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,NID from ROUTING WHERE NID='"+ROUTINGID+"')"
		SQL=SQL+" where Mother_Routing_ID='"+ROUTINGID+"'"
		rsRouting.open SQL,conn,1,3
	 
		'2)Copy Station
		rsRouting.open "SELECT STATIONS_INDEX FROM ROUTING WHERE NID='"+ROUTINGID+"'",conn,1,3
		Station_Index=split(rsRouting("STATIONS_INDEX"),",")
		for i=Lbound(Station_Index)   to   Ubound(Station_Index) 	
			SQL="SELECT * FROM ROUTING WHERE STATIONS_INDEX LIKE '%"+Station_Index(i)+"%'"
			set rsStationIsIn=server.createobject("adodb.recordset")
			rsStationIsIn.open SQL,conn,1,3
			if(rsStationIsIn.recordcount=0) then
				STATION_ID="ST"&NID_SEQ("STATION")
				SQL="INSERT INTO Station"
				SQL=SQL+"("
				SQL=SQL+"NID,STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,"
				SQL=SQL+" INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,STATION_ENTER_DEFECTCODE,MOTHER_STATION_ID,ROUTING_ID"
				SQL=SQL+")"
				SQL=SQL+"select '"+STATION_ID+"',STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,"
				SQL=SQL+" INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,"
				SQL=SQL+" WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,'"+STATION_ID+"','"+Station_Index(i)+"','"+ROUTINGID+"'"
				SQL=SQL+" FROM STATION_NEW WHERE NID='"+Station_Index(i)+"'"
				set rsStation=server.createobject("adodb.recordset")
				rsStation.open SQL,conn,1,3
				set rsStation=nothing
			else
				 set rsStationIsInPart=server.createobject("adodb.recordset")
				' SQL="UPDATE STATION SET (STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,  "
'				 SQL=SQL+"INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,STATION_ENTER_DEFECTCODE,MOTHER_STATION_ID,ROUTING_ID) "
'				 SQL=SQL+"="
'				 SQL=SQL+"(SELECT STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,"
'				 SQL=SQL+"INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,"
'				 SQL=SQL+" WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,'"+STATION_ID+"','"+Station_Index(i)+"','"+ROUTINGID+"'"
'				 SQL=SQL+" FROM STATION_NEW WHERE NID='"+Station_Index(i)+"')"
'				 SQL=SQL+" WHERE "
				 SQL="SELECT * FROM STATION WHERE MOTHER_STATION_ID='"+Station_Index(i)+"' AND ROUTING_ID='"+ROUTINGID+"'"
				 rsStationIsInPart.open SQL,conn,1,3
				 if(rsStationIsInPart.recordcount>0) then
				 	STATION_ID=rsStationIsInPart("NID")
				 else
				 	STATION_ID="ST"&NID_SEQ("STATION")
					SQL="INSERT INTO Station"
					SQL=SQL+"("
					SQL=SQL+"NID,STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,"
					SQL=SQL+" INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,STATION_ENTER_DEFECTCODE,MOTHER_STATION_ID,ROUTING_ID"
					SQL=SQL+")"
					SQL=SQL+"select '"+STATION_ID+"',STATION_NUMBER,STATION_NAME,STATION_CHINESE_NAME,FACTORY_ID,SECTION_ID,"
					SQL=SQL+" INITAIL_QUANTITY_TYPE,WIP_REPORT_COLUMN,"
					SQL=SQL+" WIP_SEQUENCY,OUTPUT_REPORT_COLUMN,OUTPUT_SEQUENCY,TRANSACTION_TYPE,'"+STATION_ID+"','"+Station_Index(i)+"','"+ROUTINGID+"'"
					SQL=SQL+" FROM STATION_NEW WHERE NID='"+Station_Index(i)+"'"
					set rsStationAdd=server.createobject("adodb.recordset")
					rsStationAdd.open SQL,conn,1,3
					set rsStationAdd=nothing
				 	
				 end if
			end if
			'update Part table
			set rsStation2=server.createobject("adodb.recordset")
			SQL="update part set STATIONS_INDEX=STATIONS_INDEX || '"+STATION_ID+",' WHERE NID='"+Part_ID+"'"
			rsStation2.open SQL,conn,1,3
			set rsStation2=nothing
			'response.write SQL+"<BR>"
			'Copy defect 
			'如果没有,新增
			SQL=" INSERT INTO DEFECTCODE"
			SQL=SQL+"(NID,DEFECT_CODE,DEFECT_NAME,DEFECT_CHINESE_NAME,FACTORY_ID,TRANSACTION_TYPE,STATION_ID,MOTHER_DEFECT_ID,ROUTING_ID)"
			SQL=SQL+ "select 'DF' || lpad(DEFECTCODE_SEQ.nextval,8,'0'), B.DEFECT_CODE,B.DEFECT_NAME,B.DEFECT_CHINESE_NAME,B.FACTORY_ID,"
			SQL=SQL+" A.DEFECTCODE_TRANSACTION_TYPE,'"+STATION_ID+"',B.NID,'"+ROUTINGID+"'"
			SQL=SQL+" from ROUTING_DEFECT_DETAIL A ,DEFECTCODE_NEW B "
			SQL=SQL+" WHERE ROUTING_ID='"+ROUTINGID+"' AND STATION_ID='"+Station_Index(i)+"'"
			SQL=SQL+" AND A.DEFECTCODE_ID=B.NID(+) AND DEFECTCODE_ID NOT IN (SELECT MOTHER_DEFECT_ID FROM DEFECTCODE WHERE STATION_ID='"+STATION_ID+"' AND ROUTING_ID='"+ROUTINGID+"')"
			set rsDefectCode=server.createobject("adodb.recordset")
			rsDefectCode.open SQL,conn,1,3
			'response.write SQL+"<BR>"
			set rsDefectCode=nothing
		 
		 	'如果存在,更新defectcode的transaction type
			SQL="SELECT DEFECTCODE_ID FROM ROUTING_DEFECT_DETAIL WHERE ROUTING_ID='"+ROUTINGID+"' AND STATION_ID='"+Station_Index(i)+"' "
			SQL=SQL+" AND DEFECTCODE_ID IN (SELECT MOTHER_DEFECT_ID FROM DEFECTCODE WHERE STATION_ID='"+STATION_ID+"' AND ROUTING_ID='"+ROUTINGID+"') "
			set rsDefectCode2=server.createobject("adodb.recordset")
			rsDefectCode2.open SQL,conn,1,3
			for mm=1 to rsDefectCode2.recordcount
				SQL=" update DEFECTCODE C "
				SQL=SQL+" set (DEFECT_CODE,DEFECT_NAME,DEFECT_CHINESE_NAME,FACTORY_ID,TRANSACTION_TYPE,STATION_ID,MOTHER_DEFECT_ID,ROUTING_ID) = "
				SQL=SQL+" ( "
				SQL=SQL+" Select B.DEFECT_CODE,B.DEFECT_NAME,B.DEFECT_CHINESE_NAME,B.FACTORY_ID,A.DEFECTCODE_TRANSACTION_TYPE,'"+STATION_ID+"',B.NID,'"+ROUTINGID+"' "
				SQL=SQL+" from ROUTING_DEFECT_DETAIL A ,DEFECTCODE_NEW B  "
				SQL=SQL+" WHERE A.ROUTING_ID='"+ROUTINGID+"' AND A.STATION_ID='"+Station_Index(i)+"' AND A.DEFECTCODE_ID=B.NID(+)" 
				SQL=SQL+" AND  DEFECTCODE_ID='"+rsDefectCode2("DEFECTCODE_ID")+"'"
				SQL=SQL+" )"
				SQL=SQL+" where C.Routing_ID='"+ROUTINGID+"' AND C.STATION_ID='"+STATION_ID+"' and C.MOTHER_DEFECT_ID='"+rsDefectCode2("DEFECTCODE_ID")+"'"
				set rsDefectCode1=server.createobject("adodb.recordset")
				rsDefectCode1.open SQL,conn,1,3
				'response.write SQL+"<BR>"
				set rsDefectCode1=nothing
				rsDefectCode2.movenext
			next 
			set rsDefectCode2=nothing
			'删除那些多余的defect code
			SQL=" DELETE FROM DEFECTCODE WHERE MOTHER_DEFECT_ID NOT IN (SELECT DEFECTCODE_ID FROM ROUTING_DEFECT_DETAIL WHERE STATION_ID='"+Station_Index(i)+"' AND ROUTING_ID='"+ROUTINGID+"' ) "
			SQL=SQL+" AND ROUTING_ID='"+ROUTINGID+"' and STATION_ID='"+STATION_ID+"' "
			set rsDefectCode3=server.createobject("adodb.recordset")
			rsDefectCode3.open SQL,conn,1,3
			'response.write SQL+"<BR>"
			set rsDefectCode3=nothing
			
			'copy action
			ACTIONID="AC"&NID_SEQ("ACTION")
			SQL=" INSERT INTO ACTION"
			SQL=SQL+"(NID,ACTION_NAME,ACTION_CHINESE_NAME,FACTORY_ID,ACTION_PURPOSE,STATION_POSITION,APPEND_ALLOW,NULL_ALLOW,WITH_LOT,"
			SQL=SQL+ " ACTION_TYPE,ELEMENT_TYPE,ELEMENT_NUMBER,STATION_ID,MOTHER_ACTION_ID ,ROUTING_ID)"
			SQL=SQL+ "select 'AC' || lpad(ACTION_SEQ.nextval,8,'0'), B.ACTION_NAME,B.ACTION_CHINESE_NAME,B.FACTORY_ID,B.ACTION_PURPOSE,B.STATION_POSITION, "
			SQL=SQL+" B.APPEND_ALLOW,B.NULL_ALLOW,B.WITH_LOT, B.ACTION_TYPE,B.ELEMENT_TYPE,B.ELEMENT_NUMBER,'"+STATION_ID+"',B.NID,'"+ROUTINGID+"'"
			SQL=SQL+" from ROUTING_ACTION_DETAIL A ,Action_NEW B  "
			SQL=SQL+" WHERE ROUTING_ID='"+ROUTINGID+"' AND STATION_ID='"+Station_Index(i)+"'"
			SQL=SQL+" AND  A.ACTION_ID=B.NID(+) AND ACTION_ID NOT IN (SELECT MOTHER_ACTION_ID FROM ACTION WHERE STATION_ID='"+STATION_ID+"' AND ROUTING_ID='"+ROUTINGID+"')"
			set rsAction=server.createobject("adodb.recordset")
			rsAction.open SQL,conn,1,3
			set rsAction=nothing
			
			'UPDATE Action Index in Station table
			'SQL="SELECT NID FROM ACTION WHERE STATION_ID='"+STATION_ID+"' and ROUTING_ID='"+ROUTINGID+"' "
			'SQL=SQL+" AND MOTHER_ACTION_ID IN (SELECT ACTION_ID FROM ROUTING_ACTION_DETAIL WHERE STATION_ID='"+Station_Index(i)+"' and ROUTING_ID='"+ROUTINGID+"' ) ORDER BY NID"
			SQL=" SELECT A.NID FROM ACTION A , ROUTING_ACTION_DETAIL RAD WHERE A.MOTHER_ACTION_ID =RAD.ACTION_ID "
			SQL=SQL+" AND A.STATION_ID='"+STATION_ID+"' and A.ROUTING_ID='"+ROUTINGID+"' AND RAD.STATION_ID='"+Station_Index(i)+"' AND RAD.ROUTING_ID='"+ROUTINGID+"'"
			SQL=SQL+" ORDER BY TO_NUMBER(RAD.ACTION_SEQENCE)"
			set rsAction1=server.createobject("adodb.recordset")
			rsAction1.open SQL,conn,1,3
			'response.write SQL
			ACTION_COUNT=0
			ACTION_INDEX_STR=""
			for J=1   to   rsAction1.recordcount	
				ACTION_COUNT=ACTION_COUNT+1
				ACTION_INDEX_STR=ACTION_INDEX_STR+rsAction1("NID")+","
				rsAction1.movenext
			NEXT 
			set rsAction1=nothing
			
			set rsAction2=server.createobject("adodb.recordset")
			SQL="UPDATE STATION SET ACTIONS_COUNT='"+cstr(ACTION_COUNT)+"',ACTIONS_Index='"+ACTION_INDEX_STR+"' WHERE NID='"+STATION_ID+"'"
			rsAction2.open  SQL,conn,1,3
			set rsAction2=nothing
		next 
		
		set rsRouting2=server.createobject("adodb.recordset")
		SQL="update part "
		'SQL=SQL+" set  stations_index=SUBSTR(stations_index,0,length(stations_index)-1)"
		SQL=SQL+" set stations_index=decode(substr(stations_index,length(stations_index),1),',',SUBSTR(stations_index,0,length(stations_index)-1),stations_index) "
		SQL=SQL+" where Mother_Routing_ID='"+ROUTINGID+"'"
		rsRouting2.open SQL,conn,1,3
		
		word="Successfully edit Routing."
		action="location.href='"&beforepath&"'"
		  'response.end
	end if
end if
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%if upcount<>0 then%>
if (confirm("There are <%=upcount%> opened jobs applied to old routine! Would you update them?"))
{location.href='EditPartUpdateJob.asp?id=<%=id%>&path=<%=path%>&query=<%=query%>'}
else
{<%=action%>;}
<%else%>
<%=action%>;
<%end if%>
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
