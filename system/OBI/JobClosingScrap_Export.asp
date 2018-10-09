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
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
rs.open SQL,conn,1,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="OBI Error for Scrap"
	sheet.cells(2,1).value="TXT ID"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	
	sheet.cells(2,2).value="Resubmittable"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	
	sheet.cells(2,3).value="Job Number"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	
	sheet.cells(2,4).value="Part Number"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
		
	sheet.cells(2,5).value="Line Name"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter

	sheet.cells(2,6).value="Scrap Quantity"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter

	sheet.cells(2,7).value="Scrap Account"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter
	
	sheet.cells(2,8).value="Scrap Reason"
	sheet.cells(2,8).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,8).Font.ColorIndex = xlColorWhite
	sheet.cells(2,8).Font.Bold = True
	sheet.cells(2,8).HorizontalAlignment = xlCenter
	
	sheet.cells(2,9).value="Last Update Time"
	sheet.cells(2,9).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,9).Font.ColorIndex = xlColorWhite
	sheet.cells(2,9).Font.Bold = True
	sheet.cells(2,9).HorizontalAlignment = xlCenter

	sheet.cells(2,10).value="ERP Job Status"
	sheet.cells(2,10).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,10).Font.ColorIndex = xlColorWhite
	sheet.cells(2,10).Font.Bold = True
	sheet.cells(2,10).HorizontalAlignment = xlCenter
	
	sheet.cells(2,11).value="ERP Error Code"
	sheet.cells(2,11).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,11).Font.ColorIndex = xlColorWhite
	sheet.cells(2,11).Font.Bold = True
	sheet.cells(2,11).HorizontalAlignment = xlCenter
	
	sheet.cells(2,12).value="ERP Error Explanation"
	sheet.cells(2,12).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,12).Font.ColorIndex = xlColorWhite
	sheet.cells(2,12).Font.Bold = True
	sheet.cells(2,12).HorizontalAlignment = xlCenter
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("TRANSACTION_ID")
	sheet.cells(i,2).value=rs("ALLOW_ERP_RESUBMIT")
	sheet.cells(i,3).value=rs("JOB_NUMBER")
	sheet.cells(i,4).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,5).value=rs("LINE_NAME")
	sheet.cells(i,6).value=rs("SCRAP_QUANTITY")
	sheet.cells(i,7).value=rs("SCRAP_ACCOUNT")
	sheet.cells(i,8).value=rs("SCRAP_REASON_ID")
	sheet.cells(i,9).value=formatdate(rs("ERP_LAST_UPDATE_TIME"),application("longdateformat"))
	sheet.cells(i,10).value=rs("ERP_JOB_STATUS")
	sheet.cells(i,11).value=rs("ERP_ERROR_CODE")
	sheet.cells(i,12).value=rs("ERP_ERROR_EXPLANATION")
i=i+1
rs.movenext
wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,12)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,12)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,12))
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

rs.close
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->