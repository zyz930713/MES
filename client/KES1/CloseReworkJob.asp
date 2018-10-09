<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
	JobNumber=session("JOB_NUMBER")
	SheetNumber=session("SHEET_NUMBER")
	JobType=session("JOB_TYPE")
	
	SQL="UPDATE JOB SET  FINISHED_STATIONS_ID=STATIONS_INDEX || ',', OPENED_STATIONS_ID=STATIONS_INDEX,STATUS='1',CLOSE_TIME=SYSDATE,"
	SQL=SQL+" CURRENT_STATION_ID=SUBSTR(STATIONS_INDEX,LENGTH(STATIONS_INDEX)-9,10),LAST_STATION_ID=SUBSTR(STATIONS_INDEX,LENGTH(STATIONS_INDEX)-9,10)"
	SQL=SQL+" WHERE JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
	rs.open SQL,conn,1,3
	
	'自动做个报废
 
	SQL="select * from JOB WHERE job_number='"+jobnumber+"' and  sheet_number='"&SheetNumber&"'"
	set rsJobInfo=server.createobject("adodb.recordset")
	rsJobInfo.open sql,conn,1,3
	if rsJobInfo.recordcount>0 then
		part_number_tag=rsJobInfo("PART_NUMBER_TAG")
		IF (session("JOB_TYPE")="C")THEN
			SQL="SELECT * FROM CHANGE_MODEL_MAPPING WHERE NEW_MODEL='"+part_number_tag+"'"
			set rsChangeModelName=server.createobject("adodb.recordset")
			rsChangeModelName.open SQL,conn,1,3
			if rsChangeModelName.recordcount>0 then
				part_number_tag=rsChangeModelName("OLD_MODEL")
			end if 
		END IF 
					
		factory=rsJobInfo("FACTORY_ID")
		line_name=rsJobInfo("LINE_NAME")
		Qty=rsJobInfo("JOB_START_QUANTITY")
		
		SQL="select * from JOB_MASTER_SCRAP WHERE NOTE='Rework Auto Scrap'"
		set rsReworkJobAutoScrap=server.createobject("adodb.recordset")
		rsReworkJobAutoScrap.open SQL,conn,1,3
		if rsReworkJobAutoScrap.recordcount=0 then
			rsReworkJobAutoScrap.addnew
			NID="CH"&NID_SEQ("PROD_SCRAP")
			rsReworkJobAutoScrap("NID")=NID
			rsReworkJobAutoScrap("JOB_NUMBER")=jobnumber
			response.write jobnumber
			response.end
			rsReworkJobAutoScrap("PART_NUMBER_TAG")=part_number_tag
			rsReworkJobAutoScrap("FACTORY_ID")=factory
			rsReworkJobAutoScrap("LINE_NAME")=line_name
			rsReworkJobAutoScrap("SCRAP_CODE")=session("CODE") & "-0000"
			rsReworkJobAutoScrap("SCRAP_TIME")=now()
			rsReworkJobAutoScrap("SCRAP_QUANTITY")=Qty
			rsReworkJobAutoScrap("NOTE")="Main Line Auto Scrap"
			rsReworkJobAutoScrap("REASON")="Main Line Auto Scrap"
			rsReworkJobAutoScrap("scrap_account")="1101"
			rsReworkJobAutoScrap("SCRAP_REASON_ID")="8"
			'rsScrap2("SCRAP_REFERENCE")=erp_refer
			rsReworkJobAutoScrap("STORE_NID")="-"
			rsReworkJobAutoScrap.update
			rsReworkJobAutoScrap.close
			
			SQL="update JOB_MASTER set FINAL_SCRAP_QUANTITY=FINAL_SCRAP_QUANTITY+"+cstr(Qty)+" where job_number='"+jobnumber+"'"
			set rsMaster=server.createobject("adodb.recordset")
			rsMaster.open SQL,conn,1,3
			
			SQL="update JOB_MASTER set STORE_STATUS='1' where job_number='"+jobnumber+"' "
			SQL=SQL+" and confirm_good_quantity+final_scrap_quantity=start_quantity"
			set rsMaster2=server.createobject("adodb.recordset")
			rsMaster2.open SQL,conn,1,3
			
			'add by jack zhang 2011-2-18
			SQL="update JOB set STATUS='1' where job_number='"+jobnumber+"' and SHEET_NUMBER='" &SheetNumber&"'"
			set rsJobInfo2=server.createobject("adodb.recordset")
			rsJobInfo2.open SQL,conn,1,3
		end if
	end if 	
	response.Redirect("ActionError.asp?alerttype=&alertmessage="&"成功关闭子工单!"+cstr(Qty)+"个Unit已经完成报废!请把实物交给Offline!")
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->