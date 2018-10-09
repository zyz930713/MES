<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
set rs1=server.CreateObject("adodb.recordset")

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
if not rs.EOF then

sheet.cells(1,1).value="Job Number"
sheet.cells(1,2).value="Part Number"
sheet.cells(1,3).value="Line"
sheet.cells(1,4).value="Station"
sheet.cells(1,5).value="Action Name"
sheet.cells(1,6).value="Action Value"
sheet.cells(1,7).value="Relative Action Name"
sheet.cells(1,8).value="Relative Action Value"

i=2
while not rs.eof
	sheet.cells(i,1).value=rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)
	sheet.Hyperlinks.Add sheet.cells(i,1),"http://kes1-barcode:8081/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"&sheetnumber="&rs("SHEET_NUMBER")&"&jobtype="&rs("JOB_TYPE")
	sheet.cells(i,2).value=rs("PART_NUMBER_TAG")
	sheet.cells(i,3).value=rs("LINE_NAME")
	sheet.cells(i,4).value=rs("STATION_NAME")
	sheet.cells(i,5).value=rs("ACTION_NAME")
	sheet.cells(i,6).value=rs("ACTION_VALUE")
	relative_action_name=""
	relative_action_value=""
	SQL1="select A.ACTION_NAME,RJA.ACTION_VALUE from JOB_ACTIONS RJA inner join ACTION A on RJA.ACTION_ID=A.NID where RJA.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and RJA.SHEET_NUMBER="&rs("SHEET_NUMBER")&" and RJA.JOB_TYPE='"&rs("JOB_TYPE")&"' and RJA.ACTION_ID=(select NID from ACTION where RELATED_ACTION_ID='"&rs("ACTION_ID")&"')"
	rs1.open SQL1,conn,1,3
	if not rs1.eof then
	relative_action_name=rs1("ACTION_NAME")
	relative_action_value=rs1("ACTION_VALUE")
	end if
	rs1.close
	sheet.cells(i,7).value=relative_action_name
	sheet.cells(i,8).value=relative_action_value

rs.movenext
i=i+1
wend

sheet.cells(i,1).value="end"
end if
rs.close
set rs1=nothing

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