<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
family_name=request.QueryString("family_name")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book2.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet
SQL=session("SQL")
'response.Write(SQL)
'response.End()
rs.open SQL,conn,1,3
	
	total_input_quantity=0
	total_output_quantity=0
	total_input_value=0
	total_output_value=0
	
	sheet.Name = "Output"
	sheet.cells(1,1).value="Daily Finance Yield Detail (Output) of "&family_name&" ("&from_time&" - "&to_time&")"
	sheet.cells(2,1).value="Part Number "
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter

	sheet.cells(2,2).value="Job Number"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	
	sheet.cells(2,3).value="Start Quantity"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	
	sheet.cells(2,4).value="Start Value"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
	sheet.cells(2,5).value="Output Quantity"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter

	sheet.cells(2,6).value="Output Value"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("PART_NUMBER")
	sheet.cells(i,2).value=rs("JOB_NUMBER")
	sheet.cells(i,3).value=rs("START_QUANTITY")
	sheet.cells(i,4).value=rs("START_VALUE")
	sheet.cells(i,5).value=rs("GOOD_QUANTITY")
	sheet.cells(i,6).value=rs("GOOD_VALUE")
	total_input_quantity=total_input_quantity+ccur(rs("START_QUANTITY"))
	total_output_quantity=total_output_quantity+ccur(rs("GOOD_QUANTITY"))
	total_input_value=total_input_value+ccur(rs("START_VALUE"))
	total_output_value=total_output_value+ccur(rs("GOOD_VALUE"))
i=i+1
rs.movenext
wend

	sheet.cells(i,1).value="Total"
    sheet.cells(i,3).value=cstr(total_input_quantity)
	sheet.cells(i,4).value=cstr(total_input_value)
	sheet.cells(i,5).value=cstr(total_output_quantity)
	sheet.cells(i,6).value=cstr(total_output_value)
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,6)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,6)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,6))
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
rs.close

SQL=session("SQL1")
'response.Write(SQL)
'response.End()
rs.open SQL,conn,1,3
OExcel.sheets(2).select
set Sheet=OExcel.activeWorkbook.ActiveSheet
	total_scarp_quantity=0
	total_scarp_value=0
	
	sheet.Name = "Scrap"
	sheet.cells(1,1).value="Daily Finance Yield Detail (Scrap) of "&family_name&" ("&from_time&" - "&to_time&")"
	sheet.cells(2,1).value="Part Number "
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter

	sheet.cells(2,2).value="Job Number"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	
	sheet.cells(2,3).value="Scrap Quantity"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	
	sheet.cells(2,4).value="Scrap Value"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("PART_NUMBER")
	sheet.cells(i,2).value=rs("JOB_NUMBER")
	sheet.cells(i,3).value=rs("QUANTITY")
	sheet.cells(i,4).value=rs("AMOUNT")
	total_scrap_quantity=total_scrap_quantity+ccur(rs("QUANTITY"))
	total_scrap_value=total_scrap_value+ccur(rs("AMOUNT"))
i=i+1
rs.movenext
wend

	sheet.cells(i,1).value="Total"
    sheet.cells(i,3).value=cstr(total_scrap_quantity)
	sheet.cells(i,4).value=cstr(total_scrap_value)
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,4)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,4)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,4))
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
rs.close

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
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->