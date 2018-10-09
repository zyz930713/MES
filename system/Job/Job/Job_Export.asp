<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
if not rs.EOF then

sheet.cells(1,1).value="No"
sheet.cells(1,2).value="Job Number"
sheet.cells(1,3).value="Part Number"
sheet.cells(1,4).value="Line"
sheet.cells(1,5).value="Progress"
sheet.cells(1,6).value="Planer"
sheet.cells(1,7).value="Start Quantity"
sheet.cells(1,8).value="Line Lost"
sheet.cells(1,9).value="Final Good Quantity"
sheet.cells(1,10).value="Final Scrap Quantity"
sheet.cells(1,11).value="Remain Quantity"
sheet.cells(1,12).value="Production Time"

i=2
while not rs.eof
	sheet.cells(i,1).value=j
	sheet.cells(i,2).value=rs("JOB_NUMBER")
	sheet.cells(i,3).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,4).value=rs("LINE_NAME")
	percent=round((clng(rs("FINAL_GOOD_QUANTITY"))+clng(rs("FINAL_SCRAP_QUANTITY")))/clng(rs("START_QUANTITY")),4)
	sheet.cells(i,5).value=formatpercent(percent,2,-1)
	sheet.cells(i,6).value=rs("CREATE_NAME")
	sheet.cells(i,7).value=rs("START_QUANTITY")
	sheet.cells(i,8).value=rs("DEFECTCODE_QUANTITY")
	sheet.cells(i,9).value=rs("FINAL_GOOD_QUANTITY")
	sheet.cells(i,10).value=rs("FINAL_SCRAP_QUANTITY")
	sheet.cells(i,11).value=clng(rs("START_QUANTITY"))-clng(rs("FINAL_GOOD_QUANTITY"))-clng(rs("FINAL_SCRAP_QUANTITY"))
	sheet.cells(i,12).value=formatdate(rs("INPUT_TIME"),application("longdateformat"))
rs.movenext
i=i+1
wend

sheet.cells(i,1).value="end"
end if
rs.close

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
	
myfile=filePath&"\"&rnd_key&".xls"

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=nothing
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