<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
where=session("where")
order=session("order")
set rsJ=server.CreateObject("adodb.recordset")
set rsU=server.CreateObject("adodb.recordset")
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")

rs.Open SQL,conn,1,3
if not rs.EOF then
t=1
while not rs.eof
	if t=1 then
	OExcel.sheets(t).select
	set sheet=OExcel.activeWorkbook.ActiveSheet
	else
	set sheet=OExcel.activeWorkbook.sheets.Add(,OExcel.sheets(t-1))
	end if
	i=0
	v=0
	sheet.name=rs("MACHINE_NUMBER")
	sheet.Tab.ColorIndex = xlColorArray(int(56*rnd))
	sheet.cells(1,1).value="Machine Number"
	sheet.cells(1,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,1).Font.ColorIndex = xlColorWhite
	sheet.cells(1,1).Font.Size = 10
	sheet.cells(1,1).Font.Bold = True
	sheet.cells(1,2).value="Job Number"
	sheet.cells(1,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,2).Font.ColorIndex = xlColorWhite
	sheet.cells(1,2).Font.Size = 10
	sheet.cells(1,2).Font.Bold = True
	sheet.cells(1,3).value="Part Number"
	sheet.cells(1,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,3).Font.ColorIndex = xlColorWhite
	sheet.cells(1,3).Font.Size = 10
	sheet.cells(1,3).Font.Bold = True
	sheet.cells(1,4).value="Start Time"
	sheet.cells(1,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,4).Font.ColorIndex = xlColorWhite
	sheet.cells(1,4).Font.Size = 10
	sheet.cells(1,4).Font.Bold = True
	
	astations=split(rs("STATIONS_USED"),",")
	v=5
	for s=0 to ubound(astations)
	sheet.cells(1,v).value=getStation(true,"TEXT",astations(s),"","","","")
	sheet.cells(1,v).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,v).Font.ColorIndex = xlColorWhite
	sheet.cells(1,v).Font.Size = 10
	sheet.cells(1,v).Font.Bold = True
	v=v+1
	sheet.cells(1,v).value="Yield"
	sheet.cells(1,v).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(1,v).Font.ColorIndex = xlColorWhite
	sheet.cells(1,v).Font.Size = 10
	sheet.cells(1,v).Font.Bold = True
	v=v+1
	next
	
	i=2
	SQLJ="select 1,J.JOB_NUMBER,J.SHEET_NUMBER,J.START_TIME,J.PART_NUMBER_TAG from JOB J inner join JOB_ACTIONS JA on J.JOB_NUMBER=JA.JOB_NUMBER and J.SHEET_NUMBER=JA.SHEET_NUMBER inner join PART P on J.PART_NUMBER_ID=P.NID where JA.ACTION_VALUE='"&rs("MACHINE_NUMBER")&"' "&where&order
	rsJ.open SQLJ,conn,1,3
	if not rsJ.eof then
	while not rsJ.eof 
	sheet.cells(i,1).value=rs("MACHINE_NUMBER")
	sheet.cells(i,1).Font.Size = 10
	sheet.cells(i,2).value=rsJ("JOB_NUMBER")&"-"&repeatstring(rsJ("SHEET_NUMBER"),"0",3)
	sheet.cells(i,2).Font.Size = 10
	sheet.cells(i,3).value=rsJ("PART_NUMBER_TAG")
	sheet.cells(i,3).Font.Size = 10
	sheet.cells(i,4).value=formatdate(rsJ("START_TIME"),application("longdateformat"))
	sheet.cells(i,4).Font.Size = 10
	v=5
		for s=0 to ubound(astations)
			SQLU="select STATION_START_QUANTITY,STATION_DEFECTCODE_QUANTITY,STATION_ASSEMBLY_YIELD from JOB_STATIONS where JOB_NUMBER='"&rsJ("JOB_NUMBER")&"' and SHEET_NUMBER='"&rsJ("SHEET_NUMBER")&"' and STATION_ID='"&astations(s)&"'"
			rsU.open SQLU,conn,1,3
			if not rsU.eof then
			sheet.cells(i,v).value=rsU("STATION_DEFECTCODE_QUANTITY")
			sheet.cells(i,v).Font.Size = 10
			v=v+1
			sheet.cells(i,v).value=formatpercent(csng(rsU("STATION_ASSEMBLY_YIELD")),2,-1)
			sheet.cells(i,v).Font.Size = 10
			v=v+1
			end if
			rsU.close
		next
	i=i+1
	rsJ.movenext
	wend
	end if
	rsJ.close
	
	session("aerror")=t&"-"&i&"-"&v
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,v-1))
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

rs.MoveNext
t=t+1
wend
end if
rs.Close

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
Response.ContentType = "application/vnd.ms-excel"
Response.BinaryWrite bytes
response.end

set rsJ=nothing
set rsU=nothing
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->