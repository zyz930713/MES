<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
set rsJ=server.CreateObject("adodb.recordset")
set rsS=server.CreateObject("adodb.recordset")
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQLJob")
'SQL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null and J.JOB_NUMBER like '%1221567%' and J.JOB_TYPE='N' and J.STATUS=1 and J.START_TIME>=to_date('2008-5-23','yyyy-mm-dd') and instr('FA00000002,FA00000001,FA00000003',P.FACTORY_ID)>0 order by J.PART_NUMBER_ID,J.JOB_NUMBER,J.SHEET_NUMBER"

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
session("aerror")=SQL
if not rs.EOF then
sheet.cells(1,1).value="Job Number"
sheet.cells(1,2).value="Part Number"
sheet.cells(1,3).value="Line"
sheet.cells(1,4).value="Quantity"
sheet.cells(1,5).value="Start Time"
sheet.cells(1,6).value="Close Time"
sheet.cells(1,7).value="Sub Job Number"
sheet.cells(1,8).value="Quantity"
sheet.cells(1,9).value="Routine"
sheet.cells(1,10).value="Station Name"
sheet.cells(1,11).value="Start Time"
sheet.cells(1,12).value="Close Time"
sheet.cells(1,13).value="Elapse Time"
sheet.cells(1,14).value="Store Quantity"
sheet.cells(1,15).value="Store Time"
master_row=2
while not rs.eof
	sub_row=master_row
	store_row=master_row
	'if rs("JOB_NUMBER")<>previous_job_number then
		SQLJ="select J.*,P.PART_NUMBER from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER='"&rs("JOB_NUMBER")&"' order by J.SHEET_NUMBER"
		rsJ.open SQLJ,conn,1,3
		if not rsJ.eof then
		while not rsJ.eof
		astation=split(rsJ("STATIONS_INDEX"),",")
		sheet.cells(sub_row,7).value=rsJ("JOB_NUMBER")&"-"&replace(rsJ("JOB_TYPE"),"N","")&repeatstring(rsJ("SHEET_NUMBER"),"0",3)
		sheet.Hyperlinks.Add sheet.cells(sub_row,7),"http://"&application("HostServer")&"/SubJob/JobDetail.asp?jobnumber="&rsJ("JOB_NUMBER")&"&sheetnumber="&rsJ("SHEET_NUMBER")&"&jobtype="&rsJ("JOB_TYPE")
		sheet.cells(sub_row,7).Font.Size = 10
		sheet.cells(sub_row,8).value=rsJ("JOB_START_QUANTITY")
		sheet.cells(sub_row,9).value=rsJ("PART_NUMBER")
			station_row=sub_row
			for j=0 to ubound(astation)
				SQLS="select JS.*,S.STATION_NAME from JOB_STATIONS JS inner join STATION S on JS.STATION_ID=S.NID where JS.JOB_NUMBER='"&rsJ("JOB_NUMBER")&"' and JS.SHEET_NUMBER='"&rsJ("SHEET_NUMBER")&"' and JS.STATION_ID='"&astation(j)&"'"
				rsS.open SQLS,conn,1,3
				if not rsS.eof then
					sheet.cells(station_row,10).value=rsS("STATION_NAME")
					sheet.cells(station_row,11).value=formatdate(rsS("START_TIME"),application("longdateformat"))
					sheet.cells(station_row,12).value=formatdate(rsS("CLOSE_TIME"),application("longdateformat"))
					sheet.cells(station_row,13).value=datediff("n",rsS("CLOSE_TIME"),rsS("START_TIME"))
				end if
				rsS.close
				station_row=station_row+1
			next
		sheet.range(sheet.cells(sub_row,7),sheet.cells(station_row-1,7)).merge
		sheet.range(sheet.cells(sub_row,8),sheet.cells(station_row-1,8)).merge
		sheet.range(sheet.cells(sub_row,9),sheet.cells(station_row-1,9)).merge
		sub_row=station_row
		rsJ.movenext
		wend
		end if
		rsJ.close
	'next
	sheet.cells(master_row,1).value=rs("JOB_NUMBER")
	sheet.Hyperlinks.Add sheet.cells(master_row,1),"http://"&application("HostServer")&"/Job/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")
	sheet.cells(master_row,1).Font.Size = 10
	sheet.cells(master_row,2).value=rs("PART_NUMBER_TAG")
	sheet.cells(master_row,3).value=rs("LINE_NAME")
	sheet.cells(master_row,4).value=rs("START_QUANTITY")
	sheet.cells(master_row,5).value=formatdate(rs("INPUT_TIME"),application("longdateformat"))
	sheet.cells(master_row,6).value=formatdate(rs("COMPLETE_DATE"),application("longdateformat"))
	SQLS="select * from JOB_MASTER_STORE where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
	rsS.open SQLS,conn,1,3
	if not rsS.eof then
	while not rsS.eof
		sheet.cells(store_row,14).value=rsS("STORE_QUANTITY")
		sheet.cells(store_row,15).value=formatdate(rsS("STORE_TIME"),application("longdateformat"))
	store_row=store_row+1
	rsS.movenext
	wend
	end if
	rsS.close
	sheet.range(sheet.cells(master_row,1),sheet.cells(sub_row-1,1)).merge
	sheet.range(sheet.cells(master_row,2),sheet.cells(sub_row-1,2)).merge
	sheet.range(sheet.cells(master_row,3),sheet.cells(sub_row-1,3)).merge
	sheet.range(sheet.cells(master_row,4),sheet.cells(sub_row-1,4)).merge
	sheet.range(sheet.cells(master_row,5),sheet.cells(sub_row-1,5)).merge
	sheet.range(sheet.cells(master_row,6),sheet.cells(sub_row-1,6)).merge
master_row=sub_row
rs.movenext
wend
end if
rs.close

set selection=sheet.range(sheet.cells(1,1),sheet.cells(master_row-1,14))
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