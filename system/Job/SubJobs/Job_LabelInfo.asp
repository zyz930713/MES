<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

SQL=session("SQL")

set rs_TotalJob=server.createobject("adodb.recordset")
rs_TotalJob.open SQL,conn,1,3

TotalSQL="select  a.job_number,a.subjobnumber, e.station_name, b.label_no,b.deviationno, B.material_part_number ,b.material_lot_number, a.amount, d.category_name "
 
TotalSQL=TotalSQL+" FROM material_count_record a, mr_dispatch b,MATERIAL_LIST C, material_category D,STATION_NEW E  "
TotalSQL=TotalSQL+" WHERE A.LABELID=B.label_no AND b.material_part_number=C.material_part_number AND C.material_type=D.category_id AND A.station_id=E.NID  AND ("
FOR i=0 to rs_TotalJob.recordcount-2
	TotalSQL=TotalSQL+"  a.JOB_NUMBER='"&rs_TotalJob("JOB_NUMBER").value&"' AND a.subjobnumber='"&rs_TotalJob("SHEET_NUMBER").value&"' or "
	rs_TotalJob.movenext
next 
TotalSQL=TotalSQL+"  a.JOB_NUMBER='"&rs_TotalJob("JOB_NUMBER").value&"' AND a.subjobnumber='"&rs_TotalJob("SHEET_NUMBER").value&"' ) "

Response.ContentType = "application/vnd.ms_excel" 
Response.AddHeader "content-Disposition","attachment;filename=Job_Info.xls"

response.write "<Table border=1 bordercolor='#009999' >"
response.write "<tr>"
response.write " <td>Job_Number</td><td>Sheet_Number</td><td>Station_Name</td><td>Label_NO</td><td>Deviation_NO</td><td>Material_Part_Number</td>"
response.write "<td>Material_Lot_Number</td>"
response.write "<td>Qty</td>"
response.write "<td>Material_Category_Name</td>"
response.write "</tr>"

rs.open TotalSQL,conn,1,3
for i=0 to rs.recordcount-1
	response.write "<tr>"
	for j=0 to rs.Fields.Count-1   	
		response.write "<td>" 
		response.write rs(j).value
		response.write "</td>" 
	next  
	rs.movenext
	response.write "</tr>"
next  
response.write "</Table>"
response.end

%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
