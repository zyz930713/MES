<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
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
rsPR.open SQL,connPR,1,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="可用型号清单"
	sheet.cells(2,1).value="型号名"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
		
i=3
if not rsPR.eof then
while not rsPR.eof
	sheet.cells(i,1).value=rsPR("segment1")
i=i+1
rsPR.movenext
wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,1)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,1)).merge
	
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

rsPR.close%>
<!--#include virtual="/WOCF/POCF_Close.asp" -->