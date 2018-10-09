<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetWIPJobQuantity.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
set rsT=server.CreateObject("adodb.recordset")
line_name=request.QueryString("line_name")
WIP_name=trim(request.QueryString("WIP_name"))
WIP_report_time=trim(request.QueryString("WIP_report_time"))
section_id=trim(request.QueryString("section_id"))
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME,WIP_INCLUDED_STATIONS from STATION where WIP_REPORT_COLUMN=1 and SECTION_ID='"&section_id&"' order by WIP_SEQUENCY"
rsT.open SQLT,conn,1,3
dim total_station_quantity
redim total_station_quantity(rsT.recordcount-1)

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="WIP Report for "&line_name&" of "&WIP_name&" on "&WIP_report_time
	sheet.cells(2,1).value="PART Number"
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
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
	if not rsT.eof then
	j=0
	While not rsT.eof
	station_quantity=getWIPJobQuantity(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rsT("NID"),rsT("WIP_INCLUDED_STATIONS"))
	sheet.cells(i,j+2).value=station_quantity
	total_station_quantity(j)=total_station_quantity(j)+csng(station_quantity)
	j=j+1
	rsT.movenext
	wend
	rsT.movefirst
	end if
i=i+1
rs.movenext
wend

	sheet.cells(i,1).value="Total"
	for k=0 to ubound(total_station_quantity)
    sheet.cells(i,k+2).value=total_station_quantity(k)
	next
	
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,j+1)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,j+1)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,j+1))
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

rsT.close
rs.close
set rsT=nothing
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->