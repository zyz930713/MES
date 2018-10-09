<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetJobSetStartQuantity.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetJobStationGoodQuantity.asp" -->
<!--#include virtual="/Functions/GetJobStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
set rsU=server.CreateObject("adodb.recordset")
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet
rs.open SQL,conn,1,3

sheet.cells(1,1).value="NO"
sheet.cells(1,2).value="Line"
sheet.cells(1,3).value="Job Number"
sheet.cells(1,4).value="Part Number"
sheet.cells(1,5).value="Start Time"
sheet.cells(1,6).value="Start Quantity"
sheet.cells(1,7).value="Total Defect Code"
sheet.cells(1,8).value="Actual Yield"

i=1
if not rs.eof then

dim thisjobnumber()
dim thissheetnumber()
dim thispartnumbertag()
dim thisstart_time()
dim thisstations_index()
dim thisjobstartquantity()
dim thisjobtotaldefectcodequantity()
dim thisjobgoodquantity()
dim thisjob_assembly_yield()
redim thisjobnumber(0)
redim thissheetnumber(0)
redim thispartnumbertag(0)
redim thisstart_time(0)
redim thisstations_index(0)
redim thisjobstartquantity(0)
redim thisjobtotaldefectcodequantity(0)
redim thisjobgoodquantity(0)
redim thisjob_assembly_yield(0)
total_jobstartquantity=0
total_jobtotaldefectcodequantity=0
total_jobgoodquantity=0
final_assembly_yield=0
current_linename=rs("LINE_NAME")
t=1
k=1
while not rs.eof
	jobstartquantity=rs("JOB_START_QUANTITY")
	jobtotaldefectcodequantity=rs("JOB_DEFECTCODE_QUANTITY")
	jobgoodquantity=rs("JOB_GOOD_QUANTITY")
	job_assembly_yield=csng(rs("JOB_ASSEMBLY_YIELD"))
	
	if rs("LINE_NAME")=current_linename then 'Get total parameters' value of all the same jobs
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisstations_index(UBound(thisstations_index)+1)
	ReDim Preserve thisjobstartquantity(UBound(thisjobstartquantity)+1)
	ReDim Preserve thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity)+1)
	ReDim Preserve thisjobgoodquantity(UBound(thisjobgoodquantity)+1)
	ReDim Preserve thisjob_assembly_yield(UBound(thisjob_assembly_yield)+1)
	total_jobstartquantity=total_jobstartquantity+cint(jobstartquantity)
	total_jobtotaldefectcodequantity=total_jobtotaldefectcodequantity+cint(jobtotaldefectcodequantity)
	total_jobgoodquantity=total_jobgoodquantity+cint(jobgoodquantity)
	else
	sheet.cells(k,1).value=t
	sheet.cells(k,2).value=current_linename
	sheet.cells(k,3).value=""
	sheet.cells(k,4).value=""
	sheet.cells(k,5).value=""
	sheet.cells(k,6).value=total_jobstartquantity
	sheet.cells(k,7).value=total_jobtotaldefectcodequantity
	if total_jobstartquantity<>0 then
	final_assembly_yield=cint(total_jobgoodquantity)/cint(total_jobstartquantity)
	else
	final_assembly_yield=0
	end if
	sheet.cells(k,8).value=formatpercent(final_assembly_yield,2)
	k=k+1
	for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
	sheet.cells(k,1).value=u+1
	sheet.cells(k,2).value=""
	sheet.cells(k,3).value=thisjobnumber(u)&"-"&repeatstring(thissheetnumber(u),"0",3)
	sheet.cells(k,4).value=thispartnumbertag(u)
	sheet.cells(k,5).value=formatdate(thisstart_time(u),application("longdateformat"))
	sheet.cells(k,6).value=thisjobstartquantity(u)
	sheet.cells(k,7).value=thisjobtotaldefectcodequantity(u)
	sheet.cells(k,8).value=formatpercent(thisjob_assembly_yield(u),2)
	k=k+1
	next
	
	redim thisjobnumber(0)
	redim thissheetnumber(0)
	redim thispartnumbertag(0)
	redim thisstart_time(0)
	redim thisstations_index(0)
	redim thisjobstartquantity(0)
	redim thisjobtotaldefectcodequantity(0)
	redim thisjobgoodquantity(0)
	redim thisjob_assembly_yield(0)
	total_jobstartquantity=0
	total_jobtotaldefectcodequantity=0
	total_jobgoodquantity=0
	final_assembly_yield=0
	thisjobnumber(UBound(thisjobnumber))=rs("JOB_NUMBER")
	thissheetnumber(UBound(thissheetnumber))=rs("SHEET_NUMBER")
	thispartnumbertag(UBound(thispartnumbertag))=rs("PART_NUMBER_TAG")
	thisstart_time(UBound(thisstart_time))=rs("START_TIME")
	thisstations_index(UBound(thisstations_index))=rs("STATIONS_INDEX")
	thisjobstartquantity(UBound(thisjobstartquantity))=jobstartquantity
	thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity))=jobtotaldefectcodequantity
	thisjobgoodquantity(UBound(thisjobgoodquantity))=jobgoodquantity
	thisjob_assembly_yield(UBound(thisjob_assembly_yield))=job_assembly_yield
	ReDim Preserve thisjobnumber(UBound(thisjobnumber)+1)
	ReDim Preserve thissheetnumber(UBound(thissheetnumber)+1)
	ReDim Preserve thispartnumbertag(UBound(thispartnumbertag)+1)
	ReDim Preserve thisstart_time(UBound(thisstart_time)+1)
	ReDim Preserve thisstations_index(UBound(thisstations_index)+1)
	ReDim Preserve thisjobstartquantity(UBound(thisjobstartquantity)+1)
	ReDim Preserve thisjobtotaldefectcodequantity(UBound(thisjobtotaldefectcodequantity)+1)
	ReDim Preserve thisjobgoodquantity(UBound(thisjobgoodquantity)+1)
	ReDim Preserve thisjob_assembly_yield(UBound(thisjob_assembly_yield)+1)
	total_jobstartquantity=total_jobstartquantity+cint(jobstartquantity)
	total_jobtotaldefectcodequantity=total_jobtotaldefectcodequantity+cint(jobtotaldefectcodequantity)
	total_jobgoodquantity=total_jobgoodquantity+cint(jobgoodquantity)
	t=t+1
	end if
i=i+1
current_linename=rs("LINE_NAME")
rs.movenext
wend

	sheet.cells(k,1).value=t
	sheet.cells(k,2).value=current_linename
	sheet.cells(k,3).value=""
	sheet.cells(k,4).value=""
	sheet.cells(k,5).value=""
	sheet.cells(k,6).value=total_jobstartquantity
	sheet.cells(k,7).value=total_jobtotaldefectcodequantity
	if total_jobstartquantity<>0 then
	final_assembly_yield=cint(total_jobgoodquantity)/cint(total_jobstartquantity)
	else
	final_assembly_yield=0
	end if
	sheet.cells(k,8).value=formatpercent(final_assembly_yield,2)
	k=k+1
	
	for u=0 to ubound(thisjobnumber)-1 'Display detail of each sheet.
	sheet.cells(k,1).value=u+1
	sheet.cells(k,2).value=""
	sheet.cells(k,3).value=thisjobnumber(u)&"-"&repeatstring(thissheetnumber(u),"0",3)
	sheet.cells(k,4).value=thispartnumbertag(u)
	sheet.cells(k,5).value=formatdate(thisstart_time(u),application("longdateformat"))
	sheet.cells(k,6).value=thisjobstartquantity(u)
	sheet.cells(k,7).value=thisjobtotaldefectcodequantity(u)
	sheet.cells(k,8).value=formatpercent(thisjob_assembly_yield(u),2)
	k=k+1
	next
sheet.cells(k,1).value="end"
end if
rsD.close
rs.Close

myfile=filePath&"\"&rnd_key&".xls"

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=nothing
set OExcel=Nothing

myfile=filePath&"\"&rnd_key&".xls"

'Create a stream object
Set Stream = Server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile myfile
bytes=Stream.Read
Stream.Close
Set Stream = Nothing

set myFs=server.createObject("scripting.FileSystemObject") 
myFs.DeleteFile filePath&"*.xls" '删除该目录下所有原先产生的临时打印文件 
set myFs=nothing 

'Output the contents of the stream object
Response.ContentType = "application/vnd.ms-excel"
Response.BinaryWrite bytes
response.end

set rsD=nothing
set rsSD=nothing
%>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->