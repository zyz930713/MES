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
sheet.cells(1,2).value="Station"
sheet.cells(1,3).value="Operator"
sheet.cells(1,4).value="Job Number"
sheet.cells(1,5).value="Part Number"
sheet.cells(1,6).value="Part Type"
sheet.cells(1,7).value="Line"
sheet.cells(1,8).value="Elapsed Time "
sheet.cells(1,9).value="Start Time"
sheet.cells(1,10).value="Close Time"

i=2
ij=1
while not rs.eof
	sheet.cells(i,1).value=ij
	sheet.cells(i,2).value=rs("STATION_NAME")&" "&rs("STATION_CHINESE_NAME")
	sheet.cells(i,3).value=rs("OPERATOR_CODE")
	sheet.cells(i,4).value=rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)
	sheet.Hyperlinks.Add sheet.cells(i,4),"http://"&application("HostServer")&"/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"&sheetnumber="&rs("SHEET_NUMBER")&"&jobtype="&rs("JOB_TYPE")
	sheet.cells(i,5).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,6).value=rs("PART_NUMBER")
	sheet.cells(i,7).value=rs("LINE_NAME")
	sheet.cells(i,8).value=datediff("n",rs("START_TIME"),rs("CLOSE_TIME"))&" m"
	sheet.cells(i,9).value=formatdate(rs("START_TIME"),application("longdateformat"))
	sheet.cells(i,10).value=formatdate(rs("CLOSE_TIME"),application("longdateformat"))

rs.movenext
i=i+1
ij=ij+1
wend

sheet.cells(i,1).value="end"
end if
rs.close

set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,10))
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