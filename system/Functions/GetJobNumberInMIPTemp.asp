<%
function getJobNumberInMIPTemp(MEMS_LOT,MEMS_PART,rnd_key)
job_numbers=""
set rsM=server.CreateObject("adodb.recordset")
SQLM="select JOB_NUMBER from MIP_DETAIL_TEMP where MEMS_LOT='"&MEMS_LOT&"' and MEMS_PART='"&MEMS_PART&"' and RND_KEY='"&rnd_key&"'"
rsM.open SQLM,conn,1,3
if not rsM.eof then
while not rsM.eof
	job_numbers=job_numbers&"<a href='/Job/JobDetail.asp?jobnumber="&rsM("JOB_NUMBER")&" target='_blank'>"&rsM("JOB_NUMBER")&"</a>, "
	rsM.movenext
wend
end if
rsM.close
if job_numbers<>"" then
job_numbers=left(job_numbers,len(job_numbers)-2)
end if
getJobNumberInMIPTemp=job_numbers
set rsM=nothing
end function
%>