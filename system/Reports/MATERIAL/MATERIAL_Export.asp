<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<% Server.ScriptTimeout=999 %>
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
	OExcel.Workbooks.open(filePath&"\Book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="PVS Testing and Material"
	sheet.cells(2,1).value="AD_ID"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	
	sheet.cells(2,2).value="Product Name"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
		
	sheet.cells(2,3).value="Line Name"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
			
	sheet.cells(2,4).value="Box Name"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	
	sheet.cells(2,5).value="Measurecount"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
			
	sheet.cells(2,6).value="AMS_Measuredatetime"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
	sheet.cells(2,7).value="Test Day"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter
	
	sheet.cells(2,8).value="2D Barcode"
	sheet.cells(2,8).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,8).Font.ColorIndex = xlColorWhite
	sheet.cells(2,8).Font.Bold = True
	sheet.cells(2,8).HorizontalAlignment = xlCenter
				
	sheet.cells(2,9).value="Year"
	sheet.cells(2,9).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,9).Font.ColorIndex = xlColorWhite
	sheet.cells(2,9).Font.Bold = True
	sheet.cells(2,9).HorizontalAlignment = xlCenter
				
	sheet.cells(2,10).value="Week"
	sheet.cells(2,10).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,10).Font.ColorIndex = xlColorWhite
	sheet.cells(2,10).Font.Bold = True
	sheet.cells(2,10).HorizontalAlignment = xlCenter
			
	sheet.cells(2,11).value="Weekday"
	sheet.cells(2,11).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,11).Font.ColorIndex = xlColorWhite
	sheet.cells(2,11).Font.Bold = True
	sheet.cells(2,11).HorizontalAlignment = xlCenter
	

	sheet.cells(2,12).value="Prod Date"
	sheet.cells(2,12).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,12).Font.ColorIndex = xlColorWhite
	sheet.cells(2,12).Font.Bold = True
	sheet.cells(2,12).HorizontalAlignment = xlCenter
	
	
	sheet.cells(2,13).value="Weeknum"
	sheet.cells(2,13).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,13).Font.ColorIndex = xlColorWhite
	sheet.cells(2,13).Font.Bold = True
	sheet.cells(2,13).HorizontalAlignment = xlCenter
	
	
	sheet.cells(2,14).value="Delta time"
	sheet.cells(2,14).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,14).Font.ColorIndex = xlColorWhite
	sheet.cells(2,14).Font.Bold = True
	sheet.cells(2,14).HorizontalAlignment = xlCenter
	
	
	sheet.cells(2,15).value="Delta day"
	sheet.cells(2,15).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,15).Font.ColorIndex = xlColorWhite
	sheet.cells(2,15).Font.Bold = True
	sheet.cells(2,15).HorizontalAlignment = xlCenter
	
	sheet.cells(2,16).value="line"
	sheet.cells(2,16).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,16).Font.ColorIndex = xlColorWhite
	sheet.cells(2,16).Font.Bold = True
	sheet.cells(2,16).HorizontalAlignment = xlCenter
	
	
	sheet.cells(2,17).value="adfail"
	sheet.cells(2,17).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,17).Font.ColorIndex = xlColorWhite
	sheet.cells(2,17).Font.Bold = True
	sheet.cells(2,17).HorizontalAlignment = xlCenter
	
	sheet.cells(2,18).value="error"
	sheet.cells(2,18).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,18).Font.ColorIndex = xlColorWhite
	sheet.cells(2,18).Font.Bold = True
	sheet.cells(2,18).HorizontalAlignment = xlCenter
	
	sheet.cells(2,19).value="criterionerror"
	sheet.cells(2,19).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,19).Font.ColorIndex = xlColorWhite
	sheet.cells(2,19).Font.Bold = True
	sheet.cells(2,19).HorizontalAlignment = xlCenter
	
	
	sheet.cells(2,20).value="job_number"
	sheet.cells(2,20).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,20).Font.ColorIndex = xlColorWhite
	sheet.cells(2,20).Font.Bold = True
	sheet.cells(2,20).HorizontalAlignment = xlCenter
	
	sheet.cells(2,21).value="sheet_number"
	sheet.cells(2,21).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,21).Font.ColorIndex = xlColorWhite
	sheet.cells(2,21).Font.Bold = True
	sheet.cells(2,21).HorizontalAlignment = xlCenter
	
	sheet.cells(2,22).value="POT"
	sheet.cells(2,22).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,22).Font.ColorIndex = xlColorWhite
	sheet.cells(2,22).Font.Bold = True
	sheet.cells(2,22).HorizontalAlignment = xlCenter
	
	sheet.cells(2,23).value="TOP"
	sheet.cells(2,23).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,23).Font.ColorIndex = xlColorWhite
	sheet.cells(2,23).Font.Bold = True
	sheet.cells(2,23).HorizontalAlignment = xlCenter
	
	
	
	sheet.cells(2,24).value="Bottom"
	sheet.cells(2,24).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,24).Font.ColorIndex = xlColorWhite
	sheet.cells(2,24).Font.Bold = True
	sheet.cells(2,24).HorizontalAlignment = xlCenter
	
	sheet.cells(2,25).value="Frame"
	sheet.cells(2,25).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,25).Font.ColorIndex = xlColorWhite
	sheet.cells(2,25).Font.Bold = True
	sheet.cells(2,25).HorizontalAlignment = xlCenter
	
	
i=3
if not rs.eof then
while not rs.eof
	sheet.cells(i,1).value=rs("ad_id")
	sheet.cells(i,2).value=rs("productname")
	sheet.cells(i,3).value=rs("linename")
	sheet.cells(i,4).value=rs("boxname")
	sheet.cells(i,5).value=rs("measurecount")
	sheet.cells(i,6).value=formatdate(rs("AMS_MEASUREDATETIME"),application("longdateformat"))
	sheet.cells(i,7).value=rs("TEST_DAY")
	sheet.cells(i,8).value=rs("serialnumber")
	sheet.cells(i,9).value=rs("year")
	sheet.cells(i,10).value=rs("week")
	sheet.cells(i,11).value=rs("weekday")
	sheet.cells(i,12).value=formatdate(rs("MEASUREDATETIME"),application("longdateformat"))
	sheet.cells(i,13).value=rs("weeknum")
	sheet.cells(i,14).value=rs("delta_time")
	sheet.cells(i,15).value=rs("delta_day")
	sheet.cells(i,16).value=rs("line")
	sheet.cells(i,17).value=rs("adfail")
	sheet.cells(i,18).value=rs("error")
	sheet.cells(i,19).value=rs("criterionerror")
	sheet.cells(i,20).value=rs("job_number")
	sheet.cells(i,21).value=rs("subjobnumber")
	sheet.cells(i,22).value=rs("POT")
	sheet.cells(i,23).value=rs("TOP")
	sheet.cells(i,24).value=rs("Bottom")
	sheet.cells(i,25).value=rs("Frame")
	
	'sheet.cells(i,8).value=formatpercent(csng(rs("YIELD")),2,-1)
	'sheet.cells(i,9).value=formatdate(rs("STORE_TIME"),application("longdateformat"))
	'sheet.cells(i,10).value=rs("STORE_TYPE")
	'sheet.cells(i,11).value=rs("SUB_JOB_NUMBERS")
	'sheet.cells(i,12).value=rs("BOM_LABEL")
i=i+1
rs.movenext
wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,25)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,25)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,25))
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
	Response.contenttype="text/csv"
    Response.AddHeader "Content-Disposition", "attachment;filename=Material.csv"
	'Response.ContentType = "application/vnd.ms-excel"
	Response.BinaryWrite bytes
	response.end

rsT.close
rs.close
set rsT=nothing%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->