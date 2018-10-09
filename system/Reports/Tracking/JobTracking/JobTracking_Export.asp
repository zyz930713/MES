<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetTempFile.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<%
schedule_name=request.QueryString("schedule_name")
filePath=server.mappath("\Reports\Excel")
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

SQL=session("SQL")
rs.Open SQL,conn,1,3
if not rs.EOF then

sheet.cells(1,1).value="Job Tracking Report of "&schedule_name
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
sheet.cells(2,3).value="Quantity"
sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,3).Font.ColorIndex = xlColorWhite
sheet.cells(2,3).Font.Bold = True
sheet.cells(2,3).HorizontalAlignment = xlCenter
sheet.cells(2,4).value="Start Time"
sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,4).Font.ColorIndex = xlColorWhite
sheet.cells(2,4).Font.Bold = True
sheet.cells(2,4).HorizontalAlignment = xlCenter
sheet.cells(2,5).value="Complete Time"
sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,5).Font.ColorIndex = xlColorWhite
sheet.cells(2,5).Font.Bold = True
sheet.cells(2,5).HorizontalAlignment = xlCenter

i=3
while Not rs.EOF
sheet.cells(i,1).value=rs("JOB_NUMBER")
sheet.cells(i,2).value=rs("PART_NUMBER")
sheet.cells(i,3).value=rs("QUANTITY")
sheet.cells(i,4).value=formatdate(rs("START_TIME"),application("longdateformat"))
sheet.cells(i,5).value=formatdate(rs("CLOSE_TIME"),application("longdateformat"))
i=i+1
rs.MoveNext
wend

	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,5)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,5)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,5))
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
rs.Close

myfile=filePath&"\"&rnd_key&".xls"

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=nothing
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->