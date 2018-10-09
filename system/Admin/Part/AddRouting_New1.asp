<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
partnumber=trim(request.Form("partnumber"))
part_rule=trim(request.Form("part_rule"))
guid=trim(request.Form("txtGuid"))
stations_index=replace(request.Form("toitem2")," ","")
SEQ="RT"&NID_SEQ("ROUTING")
ROUTINGID=SEQ
'response.write GUID

SQL="select * from ROUTING where PART_NUMBER='"&partnumber&"' AND IS_DELETE='0'"
rs.open SQL,conn,1,3
if not rs.eof then
	word="Part of "&partnumber&" has existed, please input again."
	action="history.back()"
else
	rs.addnew
	rs("NID")=SEQ
	rs("JOB_RULE")=trim(request.Form("job_rule"))
	rs("PART_NUMBER")=partnumber
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
	rs("STATIONS_INDEX")=stations_index
	rs("STATIONS_ROUTINE")=request.Form("stationroutine")
	rs("STATIONS_TRANSACTION")=getStationsTransaction_New(stations_index)
	rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
	if request.Form("initial_quantity")<>"" then
		rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
	end if
	if(request.Form("yield")=null or request.Form("yield")="") then
		rs("TARGET_YIELD")=0
	else
		rs("TARGET_YIELD")=request.Form("yield")
	end if
	rs("MEET_PRIORITY")=request.Form("priority")
	rs("ROUTINE_TYPE")=0
	rs("ROUTINE_PURPOSE")=0
	rs("IS_DELETE")=0
	rs("LASTUPDATE_PERSON")=session("code")
	rs("LASTUPDATE_TIME")=date()
	rs.update
	'rs.close
	'Check Action and DefectCode
	set rs1=server.createobject("adodb.recordset")
	set rs2=server.createobject("adodb.recordset")
	set rs3=server.createobject("adodb.recordset")
	stationString=Split(stations_index,",")
	For   i   =   Lbound(stationString)   to   Ubound(stationString) 	
		SQL_CheckA="select * from ROUTING_ACTION_DETAIL_TEMP where GUID='"&guid&"' and STATION_ID='"&stationString(i)&"'"
		rs1.open SQL_CheckA,conn,1,3
		if  rs1.recordcount=0 then
			word="Please set Stations' Action and DefectCode."
			action="location.href='/Admin/Part/AddRouting_New.asp'" 
		else
			SQL_CA="insert into ROUTING_ACTION_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE) SELECT 'RA'|| lpad(ROUTING_ACTION_SEQ.nextval,8,'0'), '"&SEQ&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,ACTION_ID,ACTION_SEQENCE FROM ROUTING_ACTION_DETAIL_TEMP where GUID='"&guid&"' and station_id in ('"&stationString(i)&"')"  
			rs3.open SQL_CA,conn,1,3
		end if
		rs1.close

		set rsCD=server.createobject("adodb.recordset")
		SQL_CD="insert into ROUTING_DEFECT_DETAIL (NID,ROUTING_ID,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE) SELECT 'RD' || lpad(ROUTING_DEFECT_SEQ.nextval,8,'0'),'"&SEQ&"' AS ROUTING_ID ,STATION_ID,STATION_SEQENCE,DEFECTCODE_ID,DEFECTCODE_SEQENCE,DEFECTCODE_TRANSACTION_TYPE FROM ROUTING_DEFECT_DETAIL_TEMP where GUID='"&guid&"' and station_id in ('"&stationString(i)&"')"  
		rsCD.open SQL_CD,conn,1,3
		'end if
		'rs2.close
	next

	set rs1=nothing
	set rs2=nothing
	set rs3=nothing
	set rsCD=nothing
	'Delete Action and DefectCode in temp
	set rs4=server.createobject("adodb.recordset")
	SQL_DA="delete from ROUTING_ACTION_DETAIL_TEMP where GUID='"&guid&"'"
	rs4.open SQL_DA,conn,1,3
	
	SQL_DD="delete from ROUTING_DEFECT_DETAIL_TEMP where GUID='"&guid&"'"
	rs4.open SQL_DD,conn,1,3
	set rs4=nothing
	
	if word="" and action="" then
		'Copy the data from new table to old table
		'1) Copy routing into Part
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
		
		word="Successfully save a New Routing."
		action="location.href='"&beforepath&"'"
		 
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
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->