<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
rnd_key=gen_key(10)
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request.QueryString("jobnumber")
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
rs.open SQL,conn,1,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Store Records"
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
			
	sheet.cells(2,4).value="Operator"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
	sheet.cells(2,5).value="Input Quantity"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
			
	sheet.cells(2,6).value="Store Quantity"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
	sheet.cells(2,7).value="Inspect Quantity"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter
	
	sheet.cells(2,8).value="Ontime Yield"
	sheet.cells(2,8).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,8).Font.ColorIndex = xlColorWhite
	sheet.cells(2,8).Font.Bold = True
	sheet.cells(2,8).HorizontalAlignment = xlCenter
				
	sheet.cells(2,9).value="Store Time"
	sheet.cells(2,9).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,9).Font.ColorIndex = xlColorWhite
	sheet.cells(2,9).Font.Bold = True
	sheet.cells(2,9).HorizontalAlignment = xlCenter
				
	sheet.cells(2,10).value="Store Type"
	sheet.cells(2,10).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,10).Font.ColorIndex = xlColorWhite
	sheet.cells(2,10).Font.Bold = True
	sheet.cells(2,10).HorizontalAlignment = xlCenter
			
	sheet.cells(2,11).value="Sub Job Numbers"
	sheet.cells(2,11).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,11).Font.ColorIndex = xlColorWhite
	sheet.cells(2,11).Font.Bold = True
	sheet.cells(2,11).HorizontalAlignment = xlCenter
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("JOB_NUMBER")
	sheet.cells(i,2).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,3).value=rs("LINE_NAME")
	sheet.cells(i,4).value=rs("STORE_CODE")
	sheet.cells(i,5).value=rs("INPUT_QUANTITY")
	sheet.cells(i,6).value=rs("STORE_QUANTITY")
	sheet.cells(i,7).value=rs("INSPECT_QUANTITY")
	
	sheet.cells(i,8).value=formatpercent(csng(rs("YIELD")),2,-1)
	sheet.cells(i,9).value=formatdate(rs("STORE_TIME"),application("longdateformat"))
	sheet.cells(i,10).value=rs("STORE_TYPE")
	'sheet.cells(i,10).value=rs("SUB_JOB_NUMBERS")
i=i+1
rs.movenext
wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,10)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,10)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,10))
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