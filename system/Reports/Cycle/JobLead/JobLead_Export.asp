<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/CompareLeadTime.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
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

sheet.cells(1,1).value="Job Number"
sheet.cells(1,2).value="Part Number"
sheet.cells(1,3).value="Start Time"
sheet.cells(1,4).value="Close Time"
sheet.cells(1,5).value="Yield"
sheet.cells(1,6).value="Cycle Time (Hour)"
sheet.cells(1,7).value="Cycle Time (Day)"
sheet.cells(1,8).value="Lead Time"
sheet.cells(1,9).value="Diff"

i=2
while Not rs.EOF
sheet.cells(i,1).value=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
sheet.cells(i,2).value=rs("PART_NUMBER_TAG")
sheet.cells(i,3).value=formatdate(rs("START_TIME"),application("longdateformat"))
sheet.cells(i,4).value=formatdate(rs("CLOSE_TIME"),application("longdateformat"))
sheet.cells(i,5).value=formatpercent(csng(rs("JOB_GOOD_QUANTITY"))/csng(rs("JOB_START_QUANTITY")),2,-1)
sheet.cells(i,6).value=formatnumber(round(csng(rs("CYCLE_TIME"))/60,2),2,-1)
sheet.cells(i,7).value=formatnumber(round(csng(rs("CYCLE_TIME"))/24/60,2),2,-1)
if csng(rs("CYCLE_TIME"))>csng(rs("LEAD_TIME")) then
sheet.cells(i,7).Font.ColorIndex = xlColorRed
end if
sheet.cells(i,8).value=formatnumber(round(csng(rs("LEAD_TIME"))/24/60,2),2,-1)
sheet.cells(i,9).value=formatnumber(round((csng(rs("CYCLE_TIME"))-csng(rs("LEAD_TIME")))/24/60,2),2,-1)
i=i+1
rs.MoveNext
wend
end if
rs.Close

myfile=filePath&"\"&rnd_key&".xls"

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=nothing
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