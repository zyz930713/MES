<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetRecordedOutput.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsT=server.CreateObject("adodb.recordset")
output_id=request.QueryString("output_id")
output_name=request.QueryString("output_name")
output_report_time=request.QueryString("output_report_time")
section_id=trim(request.QueryString("section_id"))
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL="truncate table OUTPUT_DETAIL_TEMP"
rs.open SQL,conn,3,3
SQL="select L.NID,L.LINE_NAME from LINE L where L.STATUS=1 and L.SECTION_ID='"&section_id&"' order by L.LINE_NAME"
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where OUTPUT_REPORT_COLUMN=1 and SECTION_ID='"&section_id&"' order by OUTPUT_SEQUENCY"
rsT.open SQLT,conn,1,3
Tcount=rsT.recordcount+2
dim total_station_quantity
redim total_station_quantity(rsT.recordcount-1)

	filePath=server.mappath("\Reports\Excel")
	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Output Report of "&output_name&" on "&output_report_time
	sheet.cells(2,1).value="Line"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	if not rsT.eof then
	i=2
	While not rsT.eof
	sheet.cells(2,i).value=rsT("STATION_NAME")&Chr(10)&rsT("STATION_CHINESE_NAME")
	sheet.cells(2,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,i).Font.ColorIndex = xlColorWhite
	sheet.cells(2,i).Font.Bold = True
	sheet.cells(2,i).HorizontalAlignment = xlCenter
	i=i+1
	rsT.movenext
	wend
	rsT.movefirst
	end if

i=3
dim quantity()
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("LINE_NAME")
	part_job_numbers=""
	if not rsT.eof then
		redim quantity(rsT.recordcount-1)
		k=0
		While not rsT.eof
		job_numbers=""
		quantity(k)=getRecordedOutput(output_id,rs("NID"),rsT("NID"),job_numbers)
		part_job_numbers=part_job_numbers&job_numbers&","
		k=k+1
		rsT.movenext
		wend
	rsT.movefirst
	end if
	part_job_numbers=left(part_job_numbers,len(part_job_numbers)-1)
	a_part_job_numbers=split(part_job_numbers,",")
	slim_part_job_numbers=""
	for j=0 to ubound(a_part_job_numbers)
		if instr(slim_part_job_numbers,a_part_job_numbers(j))=0 then
		slim_part_job_numbers=slim_part_job_numbers&a_part_job_numbers(j)&","
		end if
	next
	slim_part_job_numbers=left(slim_part_job_numbers,len(slim_part_job_numbers)-1)
	if not rsT.eof then
		k=0
		While not rsT.eof
		sheet.cells(i,k+2).value=quantity(k)	
		total_station_quantity(k)=total_station_quantity(k)+csng(quantity(k))
		k=k+1
		rsT.movenext
		wend
	rsT.movefirst
	end if
i=i+1
rs.movenext
wend
	sheet.cells(i,1)="Total"
	for k=0 to ubound(total_station_quantity)
    	sheet.cells(i,k+2)=total_station_quantity(k)
	next
	
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,k+1)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,k+1)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,k+1))
	selection.select
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
	selection.Columns.AutoFit
	
end if

	myfile=filePath&"\"&rnd_key&".xls"
	
	OExcel.ActiveWorkbook.saveas myfile
	OExcel.workbooks.close
	OExcel.quit 
	set sheet=nothing
	set selection=nothing
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
	
	DeleteFile filePath&"\*.xls" '删除该目录下所有原先产生的临时打印文件 
	
	'Output the contents of the stream object
	Response.ContentType = "application/vnd.ms-excel"
	Response.BinaryWrite bytes
	response.end

rsT.close
rs.close
set rsT=nothing%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->