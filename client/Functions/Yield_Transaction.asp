<!--#include virtual="/Functions/SendJMail.asp" -->
<%
function Yeild_Alert(part_number_tag,line_name,job_number,sheet_number,job_type,repeated_sequence,current_station_id)
on error resume next
set rsy=server.CreateObject("adodb.recordset")
if part_number_tag<>"" then
a_part_number_tag=split(replace(part_number_tag,"SR","SR-"),"-")
prefix=a_part_number_tag(0)
else
prefix=""
end if
SQLy="select GOOD_QUANTITY,nvl(STATION_START_QUANTITY,0) as STATION_START_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&job_number&"' and SHEET_NUMBER='"&sheet_number&"' and JOB_TYPE='"&job_type&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&repeated_sequence&"'"
rsy.open SQLy,conn,1,3
if not rsy.eof then
	good_quantity=cint(rsy("GOOD_QUANTITY"))
	start_quantity=cint(rsy("STATION_START_QUANTITY"))
	if start_quantity<>0 then
	assembly_yield=cint(rsy("GOOD_QUANTITY"))/cint(rsy("STATION_START_QUANTITY"))
	else
	assembly_yield=0
	end if
end if
rsy.close

SQLy="select distinct U.EMAIL from USERS_YIELD_ALERT UA inner join USERS U on UA.USER_CODE=U.USER_CODE where (UA.ALERT_NAME='"&prefix&"' and UA.ALERT_TYPE=1 and UA.ALERT_YIELD>="&assembly_yield&") or (UA.ALERT_NAME='"&line_name&"' and UA.ALERT_TYPE=2 and UA.ALERT_YIELD>="&assembly_yield&") or (UA.ALERT_NAME in (select NID from SERIES_GROUP where instr(INCLUDED_SYSTEM_ITEMS,'"&part_number_tag&"')>0) and UA.ALERT_TYPE=3 and UA.ALERT_YIELD>="&assembly_yield&")"
rsy.open SQLy,conn,1,3
if not rsy.eof then
	emailList=""
	while not rsy.eof
		if rsy("EMAIL")<>"" then
			emailList=emailList&rsy("EMAIL")&";"
		end if
	rsy.movenext
	wend

	TotalJobNumber=job_number&"-"&repeatstring(sheet_number,"0",3)
	subject = "FPY Alert: "&job_number&"-"&repeatstring(sheet_number,"0",3)&", "&part_number_tag&", "&formatpercent(assembly_yield,2,-1)
	mailBody ="Hi, Engineer<p>The first-past yield of <a href=""http://kes1-barcode:9999/Job/SubJobs/JobDetail.asp?jobnumber="&job_number&"&sheetnumber="&sheet_number&"&jobtype="&getMainJobNumberType(TotalJobNumber)&""">"&job_number&"-"&repeatstring(sheet_number,"0",3)&"</a>&nbsp;["&part_number_tag&"] on "&line_name&" line is lower than the threshold.<br> Yield="&formatpercent(assembly_yield,2,-1)&" ("&good_quantity&"/"&start_quantity&")<p><p>Developer Info:<br>Station ID:"&current_station_id&"<br>Start:"&start_quantity&"<br>Good:"&good_quantity&"<br>Defect:"&defectcode_quantity&"<br>Operator:"&session("code")&"<br>Client IP:"&request.ServerVariables("REMOTE_HOST")&"<p>Barcode System<br>"&now()&"<br>"&request.ServerVariables("HTTP_HOST")
	
	SendJMail application("MailSender"),emailList,subject,mailBody
	
end if
rsy.close
set rsy=nothing
end function

function getMainJobNumberType(JobNumber)
		arrJob=split(JobNumber,"-")
		NewJobNumber=""
		 if arrJob(1)="E" or arrJob(1)="R" then
			SubJob=arrJob(2)
		else
			SubJob=arrJob(1)
		end if
		if(SubJob<500) then
			getMainJobNumberType="N"
		end if 
		if(SubJob<600 AND SubJob>=500) then
			getMainJobNumberType="R"
		end if
		if(SubJob<700 AND SubJob>=600) then
			getMainJobNumberType="S"
		end if
		if(SubJob>=600) then
			getMainJobNumberType="C"
		end if
end function
%>