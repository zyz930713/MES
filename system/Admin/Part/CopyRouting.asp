<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->

<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
	SEQ="RT"&NID_SEQ("ROUTING")
	ROUTINGID=SEQ
	OLDROUTINGID=request.querystring("id")
	
	'response.write SEQ & "<BR>"
	'RESPONSE.WRITE OLDROUTINGID & "<BR>"
	'Copy the data from new table to old table
	'get Part_ID
	set rsPart=server.createobject("adodb.recordset")
	SQL="INSERT INTO ROUTING(NID,PART_NUMBER,PART_TYPE,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID,INITIAL_QUANTITY,STATIONS_COUNT,"
	SQL=SQL+ " STATIONS_INDEX,MAX_INTERVAL,STATIONS_TRANSACTION,STATIONS_ROUTINE,LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,"
	SQL=SQL+"  IS_DELETE,LASTUPDATE_PERSON,LASTUPDATE_TIME)"
	SQL=SQL+" SELECT '"+ROUTINGID+"', 'Copy-' || PART_NUMBER,PART_TYPE,JOB_RULE,PART_RULE,FACTORY_ID,SECTION_ID,INITIAL_QUANTITY,STATIONS_COUNT,"
	SQL=SQL+"  STATIONS_INDEX,MAX_INTERVAL,STATIONS_TRANSACTION,STATIONS_ROUTINE,LINES_INDEX,TARGET_YIELD,MEET_PRIORITY,ROUTINE_TYPE,ROUTINE_PURPOSE,"
	SQL=SQL+"  IS_DELETE,LASTUPDATE_PERSON,LASTUPDATE_TIME FROM ROUTING WHERE NID='"+OLDROUTINGID+"'"
	rsPart.open SQL,conn,1,3

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
	
	'COPY DEFECT CODE
	set rs1=server.createobject("adodb.recordset")
	SQL="INSERT INTO ROUTING_DEFECT_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE)"
	SQL=SQL+"SELECT 'RA'|| lpad(ROUTING_ACTION_SEQ.nextval,8,'0'), '"&ROUTINGID&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE "
	SQL=SQL+" FROM ROUTING_DEFECT_DETAIL WHERE ROUTING_ID='"&OLDROUTINGID&"'"	
	'RESPONSE.WRITE SQL +"<BR>"
	rs1.open SQL,conn,1,3
	
	'COPY ACTION
	set rs2=server.createobject("adodb.recordset")
	SQL="INSERT INTO ROUTING_ACTION_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE)"
	SQL=SQL+"SELECT 'RA'|| lpad(ROUTING_ACTION_SEQ.nextval,8,'0'), '"&ROUTINGID&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE "
	SQL=SQL+" FROM ROUTING_ACTION_DETAIL WHERE ROUTING_ID='"&OLDROUTINGID&"'"	
		'RESPONSE.WRITE SQL +"<BR>"
		'RESPONSE.END
	rs2.open SQL,conn,1,3
	
	'2)Copy Station
	rsRouting.open "SELECT STATIONS_INDEX FROM ROUTING WHERE NID='"+ROUTINGID+"'",conn,1,3
	Station_Index=split(rsRouting("STATIONS_INDEX"),",")
	for i=Lbound(Station_Index)   to   Ubound(Station_Index) 	
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
			
			'update Part table
			set rsStation2=server.createobject("adodb.recordset")
			SQL="update part set STATIONS_INDEX=STATIONS_INDEX || '"+STATION_ID+",' WHERE NID='"+Part_ID+"'"
			rsStation2.open SQL,conn,1,3
			set rsStation2=nothing
			 
			'Copy defect 
			SQL=" INSERT INTO DEFECTCODE"
			SQL=SQL+"(NID,DEFECT_CODE,DEFECT_NAME,DEFECT_CHINESE_NAME,FACTORY_ID,TRANSACTION_TYPE,STATION_ID,MOTHER_DEFECT_ID,ROUTING_ID)"
			SQL=SQL+ "select 'DF' || lpad(DEFECTCODE_SEQ.nextval,8,'0'), B.DEFECT_CODE,B.DEFECT_NAME,B.DEFECT_CHINESE_NAME,B.FACTORY_ID,"
			SQL=SQL+" A.DEFECTCODE_TRANSACTION_TYPE,'"+STATION_ID+"',B.NID,'"+ROUTINGID+"'"
			SQL=SQL+" from ROUTING_DEFECT_DETAIL A ,DEFECTCODE_NEW B "
			SQL=SQL+" WHERE ROUTING_ID='"+ROUTINGID+"' AND STATION_ID='"+Station_Index(i)+"'"
			SQL=SQL+" AND A.DEFECTCODE_ID=B.NID(+)"
			set rsDefectCode=server.createobject("adodb.recordset")
			 rsDefectCode.open SQL,conn,1,3
			'response.write sql & "<br>"
			set rsDefectCode=nothing
			 
			'copy action
			ACTIONID="AC"&NID_SEQ("ACTION")
			SQL=" INSERT INTO ACTION"
			SQL=SQL+"(NID,ACTION_NAME,ACTION_CHINESE_NAME,FACTORY_ID,ACTION_PURPOSE,STATION_POSITION,APPEND_ALLOW,NULL_ALLOW,WITH_LOT,"
			SQL=SQL+ " ACTION_TYPE,ELEMENT_TYPE,ELEMENT_NUMBER,STATION_ID,MOTHER_ACTION_ID ,ROUTING_ID)"
			SQL=SQL+ "select 'AC' || lpad(ACTION_SEQ.nextval,8,'0'), B.ACTION_NAME,B.ACTION_CHINESE_NAME,B.FACTORY_ID,B.ACTION_PURPOSE,B.STATION_POSITION, "
			SQL=SQL+" B.APPEND_ALLOW,B.NULL_ALLOW,B.WITH_LOT, B.ACTION_TYPE,B.ELEMENT_TYPE,B.ELEMENT_NUMBER,'"+STATION_ID+"',B.NID,'"+ROUTINGID+"'"
			SQL=SQL+" from ROUTING_ACTION_DETAIL A ,Action_NEW B  "
			SQL=SQL+" WHERE ROUTING_ID='"+ROUTINGID+"' AND STATION_ID='"+Station_Index(i)+"'"
			SQL=SQL+" AND  A.ACTION_ID=B.NID(+)"
			set rsAction=server.createobject("adodb.recordset")
			 rsAction.open SQL,conn,1,3	 
			'response.write sql & "<br>"
			set rsAction=nothing
			
			'UPDATE Action Index in Station table
			SQL="SELECT NID FROM ACTION WHERE STATION_ID='"+STATION_ID+"' and ROUTING_ID='"+ROUTINGID+"' ORDER BY NID"
			set rsAction1=server.createobject("adodb.recordset")
			rsAction1.open SQL,conn,1,3
			 
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
%>

<%
	 response.write ("<script>window.alert('Successfully Copy Routing.');window.location='Routing.asp';</script>")
%>
