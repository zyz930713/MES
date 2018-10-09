<%
function GetMachineStationDefectCode(jobnumber,sheetnumber,station_id,stationtotaldefectcodequantity)
totaldefectcodequantity=0
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,JD.STATION_ID,S.STATION_NAME,D.DEFECT_CODE,D.DEFECT_NAME,JD.DEFECT_QUANTITY from JOB_DEFECTCODES JD inner join STATION S on JD.STATION_ID=S.NID inner join DEFECTCODE D on JD.DEFECT_CODE_ID=D.NID where JD.JOB_NUMBER='"&jobnumber&"' and JD.SHEET_NUMBER='"&sheetnumber&"' and JD.STATION_ID='"&station_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
output="<table width='100%' border='1' align='left' cellpadding='0' cellspacing='0' bordercolorlight='#73A2EE' bordercolordark='#FFFFFF'><tr>"
	while not rsS.eof
	totaldefectcodequantity=totaldefectcodequantity+cint(rsS("DEFECT_QUANTITY"))
	output=output&"<td>"&rsS("DEFECT_CODE")&"</td><td>"&rsS("DEFECT_QUANTITY")&"</td>"
	rsS.movenext
	wend
output=output&"</tr></table>"
else
output=output&"<table width='100%' border='0' align='left' cellpadding='0' cellspacing='0' bordercolorlight='#73A2EE' bordercolordark='#FFFFFF'><tr><td>&nbsp;</td></tr></table>"
end if
GetMachineStationDefectCode=output
stationtotaldefectcodequantity=totaldefectcodequantity
rsS.close
set rsS=nothing
end function
%>