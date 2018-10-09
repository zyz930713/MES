<%
function getRecordedOutput(output_id,part_number,station_id,job_numbers)	
getRecordedOutput=0
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select JOB_NUMBERS,QUANTITY from OUTPUT_DETAIL where OUTPUT_ID='"&output_id&"' and PART_NUMBER_ID='"&part_number&"' and STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
getRecordedOutput=rsS("QUANTITY")
job_numbers=rsS("JOB_NUMBERS")
end if
rsS.close
set rsS=nothing
end function
%>