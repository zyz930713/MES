<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/Functions/GetRecordedWIP.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
set rsT=server.CreateObject("adodb.recordset")
WIP_ID=request.QueryString("WIP_ID")
WIP_NAME=trim(request.QueryString("WIP_NAME"))
factory_id=request.QueryString("factory_id")
from_day=request.QueryString("from_day")
to_day=request.QueryString("to_day")

SQL="select NID,SERIES_GROUP_NAME from SERIES_GROUP where FACTORY_ID='"&factory_id&"' order by SERIES_GROUP_NAME asc"
rs.open SQL,conn,1,3
SQLT="select NID,STATION_NAME,STATION_CHINESE_NAME from STATION where WIP_REPORT_COLUMN=1 and FACTORY_ID='"&factory_id&"' order by WIP_SEQUENCY"
rsT.open SQLT,conn,1,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value=WIP_NAME
	
	i=2
	if not rsT.eof then
  	While not rsT.eof
	sheet.cells(2,i).value=rsT("STATION_NAME")
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
	While not rs.eof	
		sheet.cells(i,1).value=rs("SERIES_GROUP_NAME")
		j=2
		While not rsT.eof	
			if from_day<>"" and to_day<>"" then
			sheet.cells(i,j).value=getRecordedFamilyWIP_WTD(rs("NID"),rsT("NID"),job_numbers,from_day,to_day)
			else
			sheet.cells(i,j).value=getRecordedFamilyWIP(WIP_ID,rs("NID"),rsT("NID"))
			end if
		j=j+1
		rsT.movenext
		wend
		rsT.movefirst
		i=i+1
	rs.movenext
	wend
	end if
rs.close
rsT.close
set rsT=nothing
	'export time
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,j-1)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,j-1)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,j-1))
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