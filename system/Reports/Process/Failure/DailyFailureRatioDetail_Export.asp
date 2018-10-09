<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
station_id=request.QueryString("station_id")
defectcode_id=request.QueryString("defectcode_id")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
SQL=session("SQL")
rs.open SQL,conn,1,3
total_input_quantity=0
total_output_quantity=0

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	if report_type="" then
	sheet.cells(1,1).value="Daily Failure Ratio Detail on "&report_day
	else
	sheet.cells(1,1).value="Weekly Failure Ratio Detail from "&from_time&" to "&to_time
	end if
	sheet.cells(2,1).value="Job Number"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter

	sheet.cells(2,2).value="Part Number"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	
	sheet.cells(2,3).value="Line Name"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	
	sheet.cells(2,4).value="Station Name"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
	sheet.cells(2,5).value="Defectcode Name"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
	
	sheet.cells(2,6).value="Input Quantity"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
	sheet.cells(2,7).value="Defectcode Quantity"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter

	sheet.cells(2,8).value="Ratio"
	sheet.cells(2,8).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,8).Font.ColorIndex = xlColorWhite
	sheet.cells(2,8).Font.Bold = True
	sheet.cells(2,8).HorizontalAlignment = xlCenter
	
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)
	sheet.cells(i,2).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,3).value=rs("LINE_NAME")
	sheet.cells(i,4).value=rs("STATION_NAME")
	sheet.cells(i,5).value=rs("DEFECT_NAME")
	sheet.cells(i,6).value=rs("JOB_START_QUANTITY")
	sheet.cells(i,7).value=rs("DEFECT_QUANTITY")
	sheet.cells(i,8).value=formatpercent(csng(rs("DEFECT_QUANTITY"))/csng(rs("JOB_START_QUANTITY")),2,-1)
	total_input_quantity=total_input_quantity+cint(rs("JOB_START_QUANTITY"))
	total_output_quantity=total_output_quantity+cint(rs("DEFECT_QUANTITY"))
i=i+1
rs.movenext
wend

	sheet.cells(i,1).value="Total"
    sheet.cells(i,4).value=total_input_quantity
	sheet.cells(i,5).value=total_output_quantity
	if total_input_quantity<>0 then
	total_yield=total_output_quantity/total_input_quantity
	else
	total_yield=0
	end if
	sheet.cells(i,8).value=formatpercent(total_yield,2,-1)
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,8)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,8)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,8))
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
	set sheet=Nothing
	set selection=nothing
	set OExcel=Nothing
	
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

rs.close
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->