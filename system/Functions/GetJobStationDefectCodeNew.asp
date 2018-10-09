<%
function GetJobStationDefectCode(jobnumber,sheetnumber,station_index)
output="<table width='100%' border='1' align='left' cellpadding='0' cellspacing='0' bordercolorlight='#73A2EE' bordercolordark='#FFFFFF'>"
astation=split(station_index,",")
stationsid=""
partid=""
factoryid=""
STATION_ENTER_DEFECTCODE=""
set rsS=server.CreateObject("adodb.recordset")
SQL = "select stations_index,PART_NUMBER_ID,FACTORY_ID from job where job_number='"&jobnumber&"' and sheet_number='"&sheetnumber&"'"
rsS.open SQL,conn,1,3
if not rsS.eof then
	if rsS("stations_index")<>"" then
	stationsid=replace(rsS("stations_index"),",","','")
	partid=rsS("PART_NUMBER_ID")
	factoryid=rsS("FACTORY_ID")
	end if
end if
if factoryid="FA00000021" then
  partwhere=" and APPLIED_PART_ID like '%"&partid&"%' "
  end if
rsS.close
SQLP="Select STATION_ENTER_DEFECTCODE from STATION where NID IN ('"&stationsid&"')"
rsS.open SQLP,conn,1,3
while not rsS.eof 
	if rsS("STATION_ENTER_DEFECTCODE")<>"" then
	STATION_ENTER_DEFECTCODE=STATION_ENTER_DEFECTCODE+replace(rsS("STATION_ENTER_DEFECTCODE"),",","','")+"','"
	end if
	rsS.movenext
wend
rsS.close
if STATION_ENTER_DEFECTCODE<>"" then
STATION_ENTER_DEFECTCODE=Mid(STATION_ENTER_DEFECTCODE,1,Len(STATION_ENTER_DEFECTCODE)-3)
else
STATION_ENTER_DEFECTCODE=0
end if
'SQLS="Select 1,DC.STATION_ID,S.STATION_NAME,S.STATION_CHINESE_NAME,DC.DEFECT_CODE,DC.DEFECT_NAME,DC.DEFECT_CHINESE_NAME,nvl(JD.DEFECT_QUANTITY,0) DEFECT_QUANTITY  from (SELECT * FROM DEFECTCODE where NID IN (select distinct nid from defectcode where station_id in ('"&STATION_ENTER_DEFECTCODE&"')"&partwhere&")) DC left JOIN (select * from JOB_DEFECTCODES WHERE  JOB_NUMBER='"&jobnumber&"' AND SHEET_NUMBER='"&sheetnumber&"') JD ON DC.NID = JD.DEFECT_CODE_ID inner join STATION S on  DC.STATION_ID=S.NID order by DC.STATION_ID,DC.DEFECT_CODE"
SQLS="Select 1,DC.STATION_ID,S.STATION_NAME,S.STATION_CHINESE_NAME,DC.DEFECT_CODE,DC.DEFECT_NAME,DC.DEFECT_CHINESE_NAME,nvl(JD.DEFECT_QUANTITY,0) DEFECT_QUANTITY from  DEFECTCODE DC LEFT JOIN (select DEFECT_QUANTITY,DEFECT_CODE_ID from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' AND SHEET_NUMBER='"&sheetnumber&"') JD on DC.NID = JD.DEFECT_CODE_ID INNER JOIN part a on dc.routing_id = a.mother_routing_id INNER JOIN STATION S ON DC.STATION_ID=S.NID where DC.station_id IN ('"&STATION_ENTER_DEFECTCODE&"')"&partwhere&" order by instr(stations_index,dc.nid),defect_code"

'response.Write(SQLS)
rsS.open SQLS,conn,1,3
if not rsS.eof then
	output=output&"<tr>"
	m=0
	check_station=rsS("STATION_ID")
	while not rsS.eof			
		before_station_name=rsS("STATION_CHINESE_NAME")	
		if rsS("STATION_ID")=check_station	then
			m=m+1
		end if
		rsS.movenext
		if not rsS.eof then
			if rsS("STATION_ID")<>check_station then			
				if m<>0 then
				output=output&"<td colspan='"&m&"' class='t-t-BookSli1'><div align='center'>"&before_station_name&"</div></td>"
				end if
				check_station=rsS("STATION_ID")	
				m=0
			end if
		else
			if m<>0 then
				output=output&"<td colspan='"&m&"' class='t-t-BookSli1'><div align='center'>"&before_station_name&"</div></td>"
			end if
		end if		
	wend
	rsS.movefirst
	output=output&"</tr><tr>"
	while not rsS.eof
		output=output&"<td valign=top><table><tr><td><font color=#FF0000>"&rsS("DEFECT_CHINESE_NAME")&"</font></td></tr><tr ><td>"&rsS("DEFECT_QUANTITY")&"</td></tr></table></td>"
		rsS.movenext
	wend
end if
GetJobStationDefectCode=output&"</tr></table>"
rsS.close
set rsS=nothing
end function
%>