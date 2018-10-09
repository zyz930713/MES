<%
function getRecordedPartWIP(part_WIP_id,part_number_tag,station_id,job_numbers)	
getRecordedPartWIP=0
job_numbers=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select QUANTITY,JOB_NUMBERS1,JOB_NUMBERS2,JOB_NUMBERS3,JOB_NUMBERS4,JOB_NUMBERS5,JOB_NUMBERS6,JOB_NUMBERS7,JOB_NUMBERS8,JOB_NUMBERS9,JOB_NUMBERS10 from PART_WIP_DETAIL where PART_WIP_ID='"&part_WIP_id&"' and PART_NUMBER_TAG='"&part_number_tag&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedPartWIP=rsS("QUANTITY")
	for wc=1 to 10
	job_numbers=job_numbers&rsS("JOB_NUMBERS"&wc)
	next
end if
rsS.close
set rsS=nothing
end function
%>