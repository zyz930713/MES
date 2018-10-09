<%
function getCalulatedOutput(formtime,totime,line_id,line_name,station_id,job_number,rnd_key)	
outputcount=0
job_number=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select JA.JOB_NUMBER,JA.SHEET_NUMBER,JA.ACTION_VALUE from JOB_STATIONS JS inner join JOB J ON JS.JOB_NUMBER=J.JOB_NUMBER and JS.SHEET_NUMBER=J.SHEET_NUMBER and JS.JOB_TYPE=J.JOB_TYPE inner join JOB_ACTIONS JA on JS.JOB_NUMBER=JA.JOB_NUMBER and JS.SHEET_NUMBER=JA.SHEET_NUMBER and JS.JOB_TYPE=JA.JOB_TYPE where J.LINE_NAME='"&line_name&"' and JS.STATION_ID='"&station_id&"' and JS.STATUS=2 and JS.CLOSE_TIME>=to_date('"&fromtime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME<=to_date('"&totime&"','yyyy-mm-dd hh24:mi:ss') and JA.STATION_ID=J.FIRST_STATION_ID and JA.ACTION_ID='AC00000097'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		if rsS("ACTION_VALUE")<>"" then
		outputcount=outputcount+csng(rsS("ACTION_VALUE"))
		thisjob_number=rsS("JOB_NUMBER")&"-"&replace(rsS("JOB_TYPE"),"N","")&repeatstring(rsS("SHEET_NUMBER"),"0",3)&","
			if instr(job_number,thisjob_number)=0 then
			job_number=job_number&thisjob_number&","
			end if
		else
		outputcount=outputcount+0
		thiserror=thiserror&"<a href='/Job/JobDetail.asp?jobnumber="&rsS("JOB_NUMBER")&"&sheetnumber="&rsS("SHEET_NUMBER")&"&jobtype="&rsS("JOB_TYPE")&"' target='_blank'>"&rsS("JOB_NUMBER")&"-"&replace(rsS("JOB_TYPE"),"N","")&repeatstring(rsS("SHEET_NUMBER"),"0",3)&"</a>; "
		end if
	rsS.movenext
	wend
end if
rsS.close
getCalulatedOutput=outputcount
if outputcount<>0 and job_number<>"" then
	job_number=left(job_number,len(job_number)-1)
	SQLS="select * from OUTPUT_DETAIL_TEMP"
	rsS.open SQLS,conn,3,3
	rsS.addnew
	rsS("LINE_ID")=line_id
	rsS("STATION_ID")=station_id
	rsS("QUANTITY")=outputcount
	rsS("CALCULATE_TIME")=now()
	rsS("RND_KEY")=rnd_key
	rsS("CREATE_CODE")=session("code")
	rsS.update
	rsS.close
	set rsS=nothing

	job_number=left(job_number,len(job_number)-1)
	Set rsODB=OraDatabase.CreateDynaset("select * from OUTPUT_DETAIL_TEMP where RND_KEY='"&rnd_key&"' and LINE_ID='"&line_id&"' and STATION_ID='"&station_id&"'",0) 
	rsODB.Dbedit
	rsODB("JOB_NUMBERS").AppendChunk (job_number)
	rsODB.Dbupdate
	rsODB.close
	set rsODB=nothing
end if
end function
%>