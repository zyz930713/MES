<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
finalPart_id=request.QueryString("finalPart_id")
finalPart_name=request.QueryString("finalPart_name")
finalPart_report_time=request.QueryString("finalPart_report_time")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
factory_id=trim(request.QueryString("factory_id"))
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
rs.open SQL,conn,3,3

	filePath=server.mappath("\Reports\Excel")
	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Final Yield Report of "&finalPart_name&" from "&from_time&" to "&to_time&" (generated on "&finalPart_report_time&")"
	sheet.cells(2,1).value="Part Name"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter

	sheet.cells(2,2).value="Assembly Input"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter

	sheet.cells(2,3).value="Assembly Output"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	
	sheet.cells(2,4).value="Assembly Yield"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
	sheet.cells(2,5).value="Final Input"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
	
	sheet.cells(2,6).value="Final Output"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
	sheet.cells(2,7).value="Final Yield"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter
	
	sheet.cells(2,8).value="Target Yield"
	sheet.cells(2,8).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,8).Font.ColorIndex = xlColorWhite
	sheet.cells(2,8).Font.Bold = True
	sheet.cells(2,8).HorizontalAlignment = xlCenter

i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("Part_NAME")
	sheet.cells(i,2).value=rs("ASSEMBLY_INPUT_QUANTITY")
	sheet.cells(i,3).value=rs("ASSEMBLY_OUTPUT_QUANTITY")
	sheet.cells(i,4).value=formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1)
	sheet.cells(i,5).value=rs("INPUT_QUANTITY")
	sheet.cells(i,6).value=rs("OUTPUT_QUANTITY")
	sheet.cells(i,7).value=formatpercent(csng(rs("FINAL_YIELD")),2,-1)
	sheet.cells(i,8).value=rs("TARGET_YIELD")&"%"
i=i+1
rs.movenext
wend
rs.close	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,8)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,8)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,8))
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
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->