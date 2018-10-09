<%
function getCalulatedPartWIP(part_number_tag,station_id,merged_stations_id,job_number,rnd_key)	
WIPcount=0
job_number=""
where=""
	if merged_stations_id<>"" then
	a_merged_stations_id=split(merged_stations_id,",")
		for p=0 to ubound(a_merged_stations_id)
		where=where&" or J.CURRENT_STATION_ID='"&a_merged_stations_id(p)&"'"
		next
	end if
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JA.JOB_NUMBER,JA.SHEET_NUMBER,JA.ACTION_VALUE from JOB_ACTIONS JA inner join JOB J on JA.JOB_NUMBER=J.JOB_NUMBER and JA.SHEET_NUMBER=J.SHEET_NUMBER and JA.JOB_TYPE=J.JOB_TYPE where J.JOB_NUMBER not like '%E' and J.STATUS=0 and J.PART_NUMBER_TAG='"&part_number_tag&"' and (J.CURRENT_STATION_ID='"&station_id&"' "&where&") and JA.STATION_ID=J.FIRST_STATION_ID and JA.ACTION_ID='AC00000097'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		if rsS("ACTION_VALUE")<>"" then
		WIPcount=WIPcount+csng(rsS("ACTION_VALUE"))
		thisjob_number=rsS("JOB_NUMBER")&"-"&replace(rsS("JOB_TYPE"),"N","")&repeatstring(rsS("SHEET_NUMBER"),"0",3)
			if instr(job_number,thisjob_number)=0 then
			job_number=job_number&thisjob_number&","
			end if
		else
		WIPcount=WIPcount+0
		thiserror=thiserror&"<a href='/Job/JobDetail.asp?jobnumber="&rsS("JOB_NUMBER")&"&sheetnumber="&rsS("SHEET_NUMBER")&"&jobtype="&rsS("JOB_TYPE")&"' target='_blank'>"&rsS("JOB_NUMBER")&"-"&replace(rsS("JOB_TYPE"),"N","")&repeatstring(rsS("SHEET_NUMBER"),"0",3)&"</a>; "
		end if
	rsS.movenext
	wend
end if
rsS.close
getCalulatedPartWIP=WIPcount
if WIPcount<>0 and job_number<>"" then
	job_number=left(job_number,len(job_number)-1)
	SQLS="select * from PART_WIP_DETAIL_TEMP"
	rsS.open SQLS,conn,3,3
	rsS.addnew
	rsS("PART_NUMBER_TAG")=part_number_tag
	rsS("STATION_ID")=station_id
	rsS("QUANTITY")=WIPcount
	we=1
	for wc=1 to 10
	this_job_numbers=mid(job_number,we,4000)
		if this_job_numbers<>"" then
		rsS("JOB_NUMBERS"&wc)=this_job_numbers
		else
		exit for
		end if
	we=we+4000
	next
	rsS("CALCULATE_TIME")=now()
	rsS("RND_KEY")=rnd_key
	rsS("CREATE_CODE")=session("code")
	rsS.update
	rsS.close
	set rsS=nothing
end if
end function
%>