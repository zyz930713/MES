<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
familylinelost_id=request.QueryString("familylinelost_id")
familylinelost_name=request.QueryString("familylinelost_name")
familylinelost_report_time=request.QueryString("familylinelost_report_time")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
factory_id=trim(request.QueryString("factory_id"))
SQL=session("overall_SQL")
rs.open SQL,conn,1,3
if not rs.eof then
factory_target_quantity=rs("FACTORY_TARGET_QUANTITY")
factory_target_amount=rs("FACTORY_TARGET_AMOUNT")
else
factory_target_quantity="0"
factory_target_amount="0"
end if
rs.close
SQL=session("SQL")
rs.open SQL,conn,3,3

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Family Line Lost Report of "&familylinelost_name&" from "&formatdate(from_time,application("longdateformat"))&" to "&formatdate(to_time,application("longdateformat"))&" (generated on "&familylinelost_report_time&")"
	sheet.cells(2,1).value="Family Name"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter

	sheet.cells(2,2).value="Assembly Input"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter

	sheet.cells(2,3).value="Lost Quantity"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
		
	sheet.cells(2,4).value="Lost Amount"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
		
	sheet.cells(2,5).value="Lost Percetage"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
	
	sheet.cells(2,6).value="Target Quantity"
	sheet.cells(2,6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,6).Font.ColorIndex = xlColorWhite
	sheet.cells(2,6).Font.Bold = True
	sheet.cells(2,6).HorizontalAlignment = xlCenter
	
	sheet.cells(2,7).value="Target Amount"
	sheet.cells(2,7).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,7).Font.ColorIndex = xlColorWhite
	sheet.cells(2,7).Font.Bold = True
	sheet.cells(2,7).HorizontalAlignment = xlCenter
	
i=3
if not rs.eof then
total_input=0
total_lost_quantity=0
total_lost_amount=0
while not rs.eof
	sheet.cells(i,1).value=rs("FAMILY_NAME")
	sheet.cells(i,2).value=rs("INPUT_QUANTITY")
	sheet.cells(i,3).value=rs("LINELOST_QUANTITY")
	sheet.cells(i,4).value=rs("LINELOST_AMOUNT")
	sheet.cells(i,5).value=formatpercent(csng(rs("LINELOST_PERCENTAGE")),2,-1)
	sheet.cells(i,6).value=rs("FAMILY_TARGET_QUANTITY")
	sheet.cells(i,7).value=rs("FAMILY_TARGET_AMOUNT")
	total_input=total_input+csng(rs("INPUT_QUANTITY"))
	total_lost_quantity=total_lost_quantity+csng(rs("LINELOST_QUANTITY"))
	total_lost_amount=total_lost_amount+csng(rs("LINELOST_AMOUNT"))
i=i+1
rs.movenext
wend
rs.close
	'overall yield
	sheet.cells(i,1).value="Overall"
	sheet.cells(i,1).Interior.ColorIndex=xlColorRose
	sheet.cells(i,2).value=total_input
	sheet.cells(i,2).Interior.ColorIndex=xlColorRose
	sheet.cells(i,3).value=total_lost_quantity
	sheet.cells(i,3).Interior.ColorIndex=xlColorRose
	sheet.cells(i,4).value=total_lost_amount
	sheet.cells(i,4).Interior.ColorIndex=xlColorRose
	sheet.cells(i,5).Interior.ColorIndex=xlColorRose
	sheet.cells(i,6).Interior.ColorIndex=xlColorRose
	sheet.cells(i,7).Interior.ColorIndex=xlColorRose
	'export time
	sheet.cells(i+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,7)).merge
	sheet.range(sheet.cells(i+1,1),sheet.cells(i+1,7)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i+1,7))
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