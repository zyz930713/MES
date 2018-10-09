<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
finalseries_id=request.QueryString("finalseries_id")
finalseries_name=request.QueryString("finalseries_name")
finalseries_report_time=request.QueryString("finalseries_report_time")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
factory_id=trim(request.QueryString("factory_id"))
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("overall_SQL")
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_yield=rs("FACTORY_TARGET_YIELD")
factory_target_firstyield=rs("FACTORY_TARGET_FIRSTYIELD")
else
factory_target_yield="0"
factory_target_firstyield="0"
end if
rs.close
SQL=session("SQL")
rs.open SQL,conn,3,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Final Yield Report of "&finalseries_name&" from "&from_time&" to "&to_time&" (generated on "&finalseries_report_time&")"
	sheet.cells(2,1).value="Series Name"
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
		
	sheet.cells(2,4).value="Stocked Output"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
		
	sheet.cells(2,5).value="First Passed Yield"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
	
	sheet.cells(2,6).value="First Target Yield"
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
	
	sheet.cells(2,9).value="Delta"
	sheet.cells(2,9).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,9).Font.ColorIndex = xlColorWhite
	sheet.cells(2,9).Font.Bold = True
	sheet.cells(2,9).HorizontalAlignment = xlCenter

i=3
if not rs.eof then
total_assembly_input=0
total_assembly_output=0
total_final_output=0
while not rs.eof
	sheet.cells(i,1).value=rs("SERIES_NAME")
	sheet.cells(i,2).value=rs("ASSEMBLY_INPUT_QUANTITY")
	sheet.cells(i,3).value=rs("ASSEMBLY_OUTPUT_QUANTITY")
	sheet.cells(i,4).value=rs("OUTPUT_QUANTITY")
	sheet.cells(i,5).value=formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1)
	sheet.cells(i,5).Interior.ColorIndex = GetYieldExcelColor(rs("ASSEMBLY_YIELD"),rs("SERIES_TARGET_FIRSTYIELD"))
	sheet.cells(i,6).value=formatpercent(csng(rs("SERIES_TARGET_FIRSTYIELD"))/100,2,-1)
	sheet.cells(i,7).value=formatpercent(csng(rs("FINAL_YIELD")),2,-1)
	sheet.cells(i,7).Interior.ColorIndex = GetYieldExcelColor(rs("FINAL_YIELD"),rs("SERIES_TARGET_YIELD"))
	sheet.cells(i,8).value=formatpercent(csng(rs("SERIES_TARGET_YIELD"))/100,2,-1)
	if csng(rs("FINAL_YIELD"))<>0 then
	delta=csng(rs("FINAL_YIELD"))-csng(rs("SERIES_TARGET_YIELD"))/100
	else
	delta=0
	end if
	sheet.cells(i,9).value=formatpercent(delta,2,-1)
	if rs("OVERALL_EXCEPTION")="0" then
	total_assembly_input=total_assembly_input+csng(rs("ASSEMBLY_INPUT_QUANTITY"))
	total_assembly_output=total_assembly_output+csng(rs("ASSEMBLY_OUTPUT_QUANTITY"))
	total_final_output=total_final_output+csng(rs("OUTPUT_QUANTITY"))
	else
	sheet.cells(i,1).Interior.ColorIndex=xlColorGray
	sheet.cells(i,2).Interior.ColorIndex=xlColorGray
	sheet.cells(i,3).Interior.ColorIndex=xlColorGray
	sheet.cells(i,4).Interior.ColorIndex=xlColorGray
	sheet.cells(i,5).Interior.ColorIndex=xlColorGray
	sheet.cells(i,6).Interior.ColorIndex=xlColorGray
	sheet.cells(i,7).Interior.ColorIndex=xlColorGray
	sheet.cells(i,8).Interior.ColorIndex=xlColorGray
	sheet.cells(i,9).Interior.ColorIndex=xlColorGray
	end if
i=i+1
rs.movenext
wend
if total_assembly_input<>0 then
total_assembly_yield=total_assembly_output/total_assembly_input
total_final_yield=total_final_output/total_assembly_input
else
total_assembly_yield=0
total_final_yield=0
end if
rs.close
	sheet.cells(i,1).value="Overall"
	sheet.cells(i,1).Interior.ColorIndex=xlColorRose
	sheet.cells(i,2).value=total_assembly_input
	sheet.cells(i,2).Interior.ColorIndex=xlColorRose
	sheet.cells(i,3).value=total_assembly_output
	sheet.cells(i,3).Interior.ColorIndex=xlColorRose
	sheet.cells(i,4).value=total_final_output
	sheet.cells(i,4).Interior.ColorIndex=xlColorRose
	sheet.cells(i,5).value=formatpercent(total_assembly_yield,2,-1)
	sheet.cells(i,5).Interior.ColorIndex=GetYieldExcelColor(total_assembly_yield,factory_target_firstyield)
	sheet.cells(i,6).value=formatpercent(csng(factory_target_firstyield)/100,2,-1)
	sheet.cells(i,6).Interior.ColorIndex=xlColorRose
	sheet.cells(i,7).value=formatpercent(total_final_yield,2,-1)
	sheet.cells(i,7).Interior.ColorIndex=GetYieldExcelColor(total_final_yield,factory_target_yield)
	sheet.cells(i,8).value=formatpercent(csng(factory_target_yield)/100,2,-1)
	sheet.cells(i,8).Interior.ColorIndex=xlColorRose
	if csng(total_final_yield)<>0 then
	delta=csng(total_final_yield)-csng(factory_target_yield)/100
	else
	delta=0
	end if
	sheet.cells(i,9).value=formatpercent(delta,2,-1)
	sheet.cells(i,9).Interior.ColorIndex=xlColorRose
	'Remind*3
	sheet.cells(i+1,1).Interior.ColorIndex=xlColorGreen
	sheet.cells(i+1,2).value="means cum yield meet target"
	sheet.range(sheet.cells(i+1,2),sheet.cells(i+1,9)).merge
	sheet.cells(i+2,1).Interior.ColorIndex=xlColorYellow
	sheet.cells(i+2,2).value="means delta is between 0% to 0.5%"
	sheet.range(sheet.cells(i+2,2),sheet.cells(i+2,9)).merge
	sheet.cells(i+3,1).Interior.ColorIndex=xlColorRed
	sheet.cells(i+3,2).value="means delta is bigger than 0.5%"
	sheet.range(sheet.cells(i+3,2),sheet.cells(i+3,9)).merge
	'export time
	sheet.cells(i+4,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,9)).merge
	sheet.range(sheet.cells(i+4,1),sheet.cells(i+4,9)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+4,9))
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