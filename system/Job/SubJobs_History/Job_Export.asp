<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%server.ScriptTimeout=99999%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
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
sheet.cells(1,2).value="Status"
sheet.cells(1,3).value="Job Number"
sheet.cells(1,4).value="Part Number"
sheet.cells(1,5).value="Part Type"
sheet.cells(1,6).value="Line"
sheet.cells(1,7).value="Start Time"
sheet.cells(1,8).value="Close Time"
sheet.cells(1,9).value="Included Stations"
sheet.cells(1,10).value="Stations' Operators "
sheet.cells(1,11).value="Current Stations"
sheet.cells(1,12).value="Stations Qty"

i=2
ij=1
while not rs.eof
stations_index=rs("STATIONS_INDEX")
	sheet.cells(i,1).value=ij
		if rs("STATUS")="0" then
		simg="Opened"
		elseif rs("STATUS")="1" then
		simg="Closed"
		elseif rs("STATUS")="2" then
		simg="Paused"
		elseif rs("STATUS")="3" then
		simg="Locked"
		elseif rs("STATUS")="4" then
		simg="Aborted"
		end if
	sheet.cells(i,2).value=simg

	sheet.cells(i,3).value=rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)
	sheet.Hyperlinks.Add sheet.cells(i,3),"http://kes1-barcode:8081/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"&sheetnumber="&rs("SHEET_NUMBER")&"&jobtype="&rs("JOB_TYPE")
	sheet.cells(i,4).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,5).value=rs("PART_NUMBER")
	sheet.cells(i,6).value=rs("LINE_NAME")
	sheet.cells(i,7).value=formatdate(rs("START_TIME"),application("longdateformat"))
	sheet.cells(i,8).value=formatdate(rs("CLOSE_TIME"),application("longdateformat"))
	sheet.cells(i,9).value=replace(getStation(true,"TEXT",""," where S.NID in('"&replace(stations_index,",","','")&"')","",stations_index," -> "),"&nbsp;"," ")
	sheet.cells(i,10).value=getStationOperator(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rs("JOB_TYPE"),replace(stations_index,",","','"),repeated_stations_sequence,stations_index," -> ")
	'add by jack Zhang 2009-7-17
	sheet.cells(i,11).value=replace(getStation(true,"TEXT",""," where S.NID ='"&rs("CURRENT_STATION_ID")&"'","",rs("CURRENT_STATION_ID"),""),"&nbsp;"," ")
	sheet.cells(i,12).value=rs("Current_Qty")

rs.movenext
i=i+1
ij=ij+1
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