<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")
StationIndexSQL=session("StationIndexSQL") 
'response.Write(StationIndexSQL)
'response.End() 

set rs_StationIndex=server.createobject("adodb.recordset")
rs_StationIndex.open StationIndexSQL,conn,1,3
if(rs_StationIndex.recordcount>1) then
	response.write "<script>window.alert('The jobs you selected have different stations. System cannot generate report.');window.close();</script>"
	response.end
end if
set rs_TotalJob=server.createobject("adodb.recordset")
rs_TotalJob.open SQL,conn,1,3
dim FirstJobNumber
dim FirstJobSheetNumber
IF(rs_TotalJob.recordcount<>0) then
	FirstJobNumber=rs_TotalJob("JOB_NUMBER").value
	FirstJobSheetNumber=rs_TotalJob("SHEET_NUMBER").value
end if
if(FirstJobNumber="" OR FirstJobSheetNumber="") THEN
	response.write "<script>window.alert('Please search job first!');window.close();</script>"
	response.end
END IF

SQL="select STATIONS_INDEX from  job A  WHERE JOB_NUMBER='" &FirstJobNumber& "' AND SHEET_NUMBER='" &FirstJobSheetNumber& "'"
rs.open SQL,conn,1,3
dim ArrStation
dim i,j
dim TotalSQL
ArrStation=split(rs(0).value,",")
rs.close
TotalSQL="select  distinct A.JOB_NUMBER,A.SHEET_NUMBER,B.PART_NUMBER_TAG,B.LINE_NAME,B.CLOSE_TIME,ROUND(JOB_GOOD_QUANTITY/JOB_START_QUANTITY,2),"
for i=0 to ubound(ArrStation)
	SQL="SELECT ACTIONS_INDEX FROM STATION WHERE NID='" &ArrStation(i)& "'"
	rs.open SQL,conn,1,3
	dim ArrAction
	
	TotalSQL=TotalSQL+" (SELECT OPERATOR_CODE FROM job_stations WHERE JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND STATION_ID='" & ArrStation(i) &"'),"
	ArrAction=split(rs(0).value,",")
	for j=0 to ubound(ArrAction)
		TotalSQL=TotalSQL+" (SELECT ACTION_VALUE FROM JOB_ACTIONS WHERE JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND STATION_ID='"&ArrStation(i)&"' AND ACTION_ID='"&ArrAction(j)&"'),"
	next 
	rs.close
	
'	LabelQtySQL=" select distinct bb.label_no from material_count_record aa, mr_dispatch bb,MATERIAL_LIST cc, material_category dd,STATION_NEW ee "
'	LabelQtySQL=LabelQtySQL+" WHERE AA.LABELID=BB.label_no AND bb.material_part_number=cc.material_part_number"
'	LabelQtySQL=LabelQtySQL+" AND cc.material_type=dd.category_id AND aa.station_id=ee.NID AND aa.JOB_NUMBER='"& FirstJobNumber & "' and aa.subjobnumber='"& FirstJobSheetNumber &"' "
'	LabelQtySQL=LabelQtySQL+"  and ee.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"')  "
'	set rs_TotalLabelInOneStation=server.createobject("adodb.recordset")
'	rs_TotalLabelInOneStation.open LabelQtySQL,conn,1,3
'	for j=0 to rs_TotalLabelInOneStation.recordcount-1
'		PartNumberSQL=" select bb.material_part_number from material_count_record aa, mr_dispatch bb,MATERIAL_LIST cc, material_category dd,STATION_NEW ee "
'		PartNumberSQL=PartNumberSQL+" WHERE AA.LABELID=BB.label_no AND bb.material_part_number=cc.material_part_number"
'		PartNumberSQL=PartNumberSQL+" AND cc.material_type=dd.category_id AND aa.station_id=ee.NID AND aa.JOB_NUMBER=A.JOB_NUMBER and aa.subjobnumber=A.SHEET_NUMBER "
'		PartNumberSQL=PartNumberSQL+"  and ee.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"')  and bb.label_no='"+rs_TotalLabelInOneStation(0).value+"'"
'		
'		MaterialLotNumberSQL=" select bb.material_lot_number from material_count_record aa, mr_dispatch bb,MATERIAL_LIST cc, material_category dd,STATION_NEW ee "
'		MaterialLotNumberSQL=MaterialLotNumberSQL+" WHERE AA.LABELID=BB.label_no AND bb.material_part_number=cc.material_part_number"
'		MaterialLotNumberSQL=MaterialLotNumberSQL+" AND cc.material_type=dd.category_id AND aa.station_id=ee.NID AND aa.JOB_NUMBER=A.JOB_NUMBER and aa.subjobnumber=A.SHEET_NUMBER "
'		MaterialLotNumberSQL=MaterialLotNumberSQL+"  and ee.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"') and bb.label_no='"+rs_TotalLabelInOneStation(0).value+"'"
'	
'		QtySQL=" select aa.amount  from material_count_record aa, mr_dispatch bb,MATERIAL_LIST cc, material_category dd,STATION_NEW ee "
'		QtySQL=QtySQL+" WHERE AA.LABELID=BB.label_no AND bb.material_part_number=cc.material_part_number"
'		QtySQL=QtySQL+" AND cc.material_type=dd.category_id AND aa.station_id=ee.NID AND aa.JOB_NUMBER=A.JOB_NUMBER and aa.subjobnumber=A.SHEET_NUMBER "
'		QtySQL=QtySQL+"  and ee.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"')  and bb.label_no='"+rs_TotalLabelInOneStation(0).value+"'"
'	
'		TotalSQL=TotalSQL+" ("+PartNumberSQL+"),"
'		TotalSQL=TotalSQL+" ("+MaterialLotNumberSQL+"),"
'		TotalSQL=TotalSQL+" ("+QtySQL+"),"
'		rs_TotalLabelInOneStation.movenext
'	next 

	SQLLabel="select B.material_part_number ,b.material_lot_number, a.amount, d.category_name from material_count_record a, mr_dispatch b,MATERIAL_LIST C, material_category D,STATION_NEW E "
	SQLLabel=SQLLabel+" WHERE A.LABELID=B.label_no AND b.material_part_number=C.material_part_number"
	SQLLabel=SQLLabel+" AND C.material_type=D.category_id AND A.station_id=E.NID AND A.JOB_NUMBER='" &rs_TotalJob("JOB_NUMBER").value& "' and A.subjobnumber='" &rs_TotalJob("SHEET_NUMBER").value& "'"
	SQLLabel=SQLLabel+"  and e.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"') order by material_part_number "
 
	set rs_LabelInfo=server.createobject("adodb.recordset")
	rs_LabelInfo.open SQLLabel,conn,1,3
	for j=0 to rs_LabelInfo.recordcount-1 
		TotalSQL=TotalSQL+"'"&rs_LabelInfo(0).value&"',"+"'"&rs_LabelInfo(1).value&"',"+"'"&rs_LabelInfo(2).value&"',"
		rs_LabelInfo.movenext
	next 
		
next 
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00010'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00012'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00018'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00019'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00031'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00032'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00033'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00034'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00035'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00036'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00037'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00038'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00039'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00040'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00041'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00042'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00043'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00044'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00045'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00046'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00047'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00048'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00049'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00050'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00051'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00052'),"
TotalSQL=TotalSQL+"( SELECT DEFECT_QUANTITY from JOB_DEFECTCODES E,DEFECTCODE F WHERE E.DEFECT_CODE_ID=F.NID AND JOB_NUMBER=A.JOB_NUMBER AND SHEET_NUMBER=A.SHEET_NUMBER AND F.DEFECT_CODE='DC00053')"
TotalSQL=TotalSQL+" FROM JOB_ACTIONS A,JOB B"
TotalSQL=TotalSQL+" WHERE A.JOB_NUMBER=B.JOB_NUMBER AND A.SHEET_NUMBER=B.SHEET_NUMBER AND ("
FOR i=0 to rs_TotalJob.recordcount-2
	TotalSQL=TotalSQL+"  a.JOB_NUMBER='"&rs_TotalJob("JOB_NUMBER").value&"' AND b.SHEET_NUMBER='"&rs_TotalJob("SHEET_NUMBER").value&"' or "
	rs_TotalJob.movenext
next 
TotalSQL=TotalSQL+"  a.JOB_NUMBER='"&rs_TotalJob("JOB_NUMBER").value&"' AND b.SHEET_NUMBER='"&rs_TotalJob("SHEET_NUMBER").value&"' )  order by a.JOB_NUMBER, a.SHEET_NUMBER "

Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=Job_Info.xls"

response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
response.write " <td>Job_Number</td><td>Sheet_Number</td><td>Part_Number</td><td>Line_Name</td><td>Close_Time</td><td>yield</td>"
for i=0 to ubound(ArrStation)
	SQL="SELECT ACTIONS_INDEX FROM STATION WHERE NID='" &ArrStation(i)& "'"
	rs.open SQL,conn,1,3
	dim ArrAction_2
	response.write "<td>"& Get_StationName(ArrStation(i)) &"OP</td>"
	ArrAction_2=split(rs(0).value,",")
	for j=0 to ubound(ArrAction_2)
		response.write "<td>"& Get_StationName(ArrStation(i)) &"("& Get_ActionName(ArrAction_2(j)) &")</td>"
	next 
	rs.close
	
	SQLLabel="select B.material_part_number ,b.material_lot_number, a.amount, d.category_name from material_count_record a, mr_dispatch b,MATERIAL_LIST C, material_category D,STATION_NEW E "
	SQLLabel=SQLLabel+" WHERE A.LABELID=B.label_no AND b.material_part_number=C.material_part_number"
	SQLLabel=SQLLabel+" AND C.material_type=D.category_id AND A.station_id=E.NID AND A.JOB_NUMBER='" &FirstJobNumber& "' and A.subjobnumber='" &FirstJobSheetNumber& "'"
	SQLLabel=SQLLabel+"  and e.nid=(select mother_station_id from station where nid='" & ArrStation(i) &"') order by material_part_number"
	set rs_LabelInfo2=server.createobject("adodb.recordset")
	rs_LabelInfo2.open SQLLabel,conn,1,3
	for j=1 to rs_LabelInfo2.recordcount 
		response.write "<td>" & rs_LabelInfo2(3).value &"(material_part_number)"& +cstr(j)+"</td>"
		response.write "<td>" & rs_LabelInfo2(3).value &"(material_lot_number)"& +cstr(j)+"</td>"
		response.write "<td>" & rs_LabelInfo2(3).value &"(qty)"& +cstr(j)+"</td>"	
		rs_LabelInfo2.movenext
	next 
	
next
response.write "<td>figure reject-refer.</td>"
response.write "<td>Irregular pattern</td>"
response.write "<td>No Response</td>"
response.write "<td>no response-refer</td>"
response.write "<td>Dist</td>"
response.write "<td>LLF</td>"
response.write "<td>SFRQLOW</td>"
response.write "<td>MPO</td>"
response.write "<td>V/B</td>"
response.write "<td>LHF</td>"
response.write "<td>GND</td>"
response.write "<td>OP</td>"
response.write "<td>GAP</td>"
response.write "<td>SF</td>"
response.write "<td>EUL</td>"
response.write "<td>IRR</td>"
response.write "<td>NRES</td>"
response.write "<td>MECH</td>"
response.write "<td>HI</td>"
response.write "<td>C/L</td>"
response.write "<td>M/R</td>"
response.write "<td>IMP</td>"
response.write "<td>L/T</td>"
response.write "<td>V/T</td>"
response.write "<td>BZ</td>"
response.write "<td>PH</td>"
response.write "<td>Others</td>"
response.write "</tr>"

rs.open TotalSQL,conn,1,3
for i=0 to rs.recordcount-1
	response.write "<tr>"
	for j=0 to rs.Fields.Count-1   	
		response.write "<td>" 
		 if(j=4) then
		 	IF(not isnull(rs(j)) or rs(j).value<>"") then
				response.write year(cdate(rs(j).value)) &"-"&month(cdate(rs(j).value)) &"-"&day(cdate(rs(j).value)) 
			else
				response.write "&nbsp;"
			end if
		 end if
		
		if(j=5) then
			response.write cstr(round(cdbl(rs(j).value)*100,2)) & "%"
		end if
		 if(j<>4 and j<> 5) then
			if isnull(rs(j))=true or rs(j).value="" then
				response.write "0"
			else
				response.write rs(j).value
			end if 
		end if
		response.write "</td>" 
	next  
	rs.movenext
	response.write "</tr>"
next  
response.write "</Table>"
response.end

function Get_StationName(Station_ID)
	set rs_Station_Name=server.createobject("adodb.recordset")
	dim SQL_Temp
	SQL_Temp="SELECT station_name FROM STATION WHERE NID='" &Station_ID& "'"
	rs_Station_Name.open SQL_Temp,conn,1,3
	if rs_Station_Name.recordcount<>0 then
		Get_StationName=rs_Station_Name(0).value
	else
		Get_StationName=""
	end if
end function 

function Get_ActionName(Action_ID)
	set rs_Action_Name=server.createobject("adodb.recordset")
	dim SQL_Temp
	SQL_Temp="SELECT action_name FROM ACTION WHERE NID='" &Action_ID& "'"
	rs_Action_Name.open SQL_Temp,conn,1,3
	if rs_Action_Name.recordcount<>0 then
		Get_ActionName=rs_Action_Name(0).value
	else
		Get_ActionName=""
	end if
end function 



%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
