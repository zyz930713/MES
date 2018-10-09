<%
function GetJobStationDefectCode(jobnumber,sheetnumber,station_index)
output="<table width='100%' border='1' align='left' cellpadding='0' cellspacing='0' bordercolorlight='#73A2EE' bordercolordark='#FFFFFF'>"
astation=split(station_index,",")
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select JD.STATION_ID,S.STATION_NAME,S.STATION_CHINESE_NAME,D.DEFECT_CODE,D.DEFECT_NAME,D.DEFECT_CHINESE_NAME,JD.DEFECT_QUANTITY from JOB_DEFECTCODES JD inner join STATION S on JD.STATION_ID=S.NID inner join DEFECTCODE D on JD.DEFECT_CODE_ID=D.NID where JD.JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' order by JD.STATION_ID"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	output=output&"<tr>"
	m=0
	for j=0 to ubound(astation)
		while not rsS.eof
			if rsS("STATION_ID")=astation(j) then
			before_station_name=rsS("STATION_CHINESE_NAME")
			m=m+1
			end if
		rsS.movenext
		wend
		if m<>0 then
		output=output&"<td colspan='"&m*2&"' class='t-b-blue'><div align='center'>"&before_station_name&"</div></td>"
		end if
	m=0
	rsS.movefirst
	next
	output=output&"</tr><tr>"
	for k=0 to ubound(astation)
		while not rsS.eof
			if rsS("STATION_ID")=astation(k) then
			output=output&"<td>"&rsS("DEFECT_CHINESE_NAME")&"</td><td>"&rsS("DEFECT_QUANTITY")&"</td>"
			end if
		rsS.movenext
		wend
	rsS.movefirst
	next
end if
GetJobStationDefectCode=output&"</tr></table>"
rsS.close
set rsS=nothing
end function
%>