<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
set rsD=server.CreateObject("adodb.recordset")
set rsSD=server.CreateObject("adodb.recordset")
SQLD="select NID,DEFECT_NAME from DEFECTCODE"
rsD.open SQLD,conn,1,3
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
if not rs.EOF then

sheet.cells(1,1).value="Job Number"
sheet.cells(1,2).value="Part Number"
sheet.cells(1,3).value="Start Time"
sheet.cells(1,4).value="Start Quantity"
sheet.cells(1,5).value="Total Defect Code"
sheet.cells(1,6).value="Yield"
d=7
while not rsD.eof
sheet.cells(1,d).value=rsD("DEFECT_NAME")
d=d+1
rsD.movenext
wend
rsD.movefirst

i=2
while Not rs.EOF
sheet.cells(i,1).value=rs("JOB_NUMBER")
sheet.cells(i,2).value=rs("PART_NUMBER")
sheet.cells(i,3).value=formatdate(rs("START_TIME"),application("longdateformat"))
sheet.cells(i,4).value=rs("JOB_START_QUANTITY")
sheet.cells(i,5).value=rs("JOB_DEFECTCODE_QUANTITY")
sheet.cells(i,6).value=formatnumber(rs("JOB_ASSEMBLY_YIELD"),10)
	d=7
	while not rsD.eof
		SQLSD="select DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&rs("JOB_NUMBER")&"' and DEFECT_CODE_ID='"&rsD("NID")&"'"
		rsSD.open SQLSD,conn,1,3
		if not rsSD.eof then
		sheet.cells(i,d).value=rsSD("DEFECT_QUANTITY")
		else
		sheet.cells(i,d).value=0
		end if
		rsSD.close
	d=d+1
	rsD.movenext
	wend
	rsD.movefirst
i=i+1
rs.MoveNext
wend

sheet.cells(i,1).value="end"
end if
rsD.close
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