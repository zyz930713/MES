<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
set rsS=server.CreateObject("adodb.recordset")
set rsA=server.CreateObject("adodb.recordset")
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
SQL=session("SQLStation")
session("aerror")=SQL
'SQL="select J.*,P.PART_NUMBER,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null and J.JOB_NUMBER like '%1221567%' and J.JOB_TYPE='N' and J.STATUS=1 and J.START_TIME>=to_date('2008-5-23','yyyy-mm-dd') and instr('FA00000002,FA00000001,FA00000003',P.FACTORY_ID)>0 order by J.PART_NUMBER_ID,J.JOB_NUMBER,J.SHEET_NUMBER"

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
if not rs.EOF then
i=1
p=1
sheet.cells(i,1).value="Job Number"
sheet.cells(i,2).value="Part Number"
sheet.cells(i,3).value="Line"
sheet.cells(i,4).value="Start Time"
sheet.cells(i,5).value="Close Time"
sheet.cells(i,6).value="Yield"
previous_part_number=rs("PART_NUMBER_ID")
while not rs.eof
	astation=split(rs("STATIONS_INDEX"),",")
	m=7
	n=m+4
	if rs("PART_NUMBER_ID")<>previous_part_number then
	i=i+2
	end if
	for j=0 to ubound(astation)
		SQLS="select JS.START_TIME,JS.CLOSE_TIME,JS.OPERATOR_CODE,S.STATION_NAME,S.STATION_CHINESE_NAME,S.ACTIONS_INDEX from JOB_STATIONS JS inner join STATION S on JS.STATION_ID=S.NID where JS.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and JS.SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and JS.STATION_ID='"&astation(j)&"'"
		rsS.open SQLS,conn,1,3
		if not rsS.eof then
		aaction=split(rsS("ACTIONS_INDEX"),",")
		if i=1 or rs("PART_NUMBER_ID")<>previous_part_number then
		sheet.cells(i,m).value=rsS("STATION_NAME")&" ("&rsS("STATION_CHINESE_NAME")&")"
		sheet.range(sheet.cells(i,m),sheet.cells(i,m+ubound(aaction)+4)).merge
		end if
			for k=0 to ubound(aaction)
				SQLA="select JA.ACTION_VALUE,A.ACTION_NAME,A.ACTION_CHINESE_NAME from JOB_ACTIONS JA inner join ACTION A on JA.ACTION_ID=A.NID where JA.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and JA.SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and JA.STATION_ID='"&astation(j)&"' and JA.ACTION_ID='"&aaction(k)&"'"
				rsA.open SQLA,conn,1,3
				if not rsA.eof then
					if i=1 or rs("PART_NUMBER_ID")<>previous_part_number then
					sheet.cells(i+1,n).value=rsA("ACTION_NAME")
					sheet.cells(i+2,n).value=rsA("ACTION_VALUE")
					p=i+2
					else
					sheet.cells(i,n).value=rsA("ACTION_VALUE")
					p=i
					end if
				end if
				rsA.close
				n=n+1
			next
			sheet.cells(p,m).value=rsS("OPERATOR_CODE")
			sheet.cells(p,m+1).value=formatdate(rsS("START_TIME"),application("longdateformat"))
			sheet.cells(p,m+2).value=formatdate(rsS("CLOSE_TIME"),application("longdateformat"))
			sheet.cells(p,m+3).value=datediff("n",rsS("START_TIME"),rsS("CLOSE_TIME"))
			m=m+ubound(aaction)+5
			n=m+4
		end if
		rsS.close
	next
	sheet.cells(p,1).value=rs("JOB_NUMBER")&"-"&replace(rs("JOB_TYPE"),"N","")&repeatstring(rs("SHEET_NUMBER"),"0",3)
	sheet.Hyperlinks.Add sheet.cells(p,1),"http://"&application("HostServer")&"/Job/SubJobs/JobDetail.asp?jobnumber="&rs("JOB_NUMBER")&"&sheetnumber="&rs("SHEET_NUMBER")&"&jobtype="&rs("JOB_TYPE")
	sheet.cells(p,1).Font.Size = 10
	sheet.cells(p,2).value=rs("PART_NUMBER_TAG")
	sheet.cells(p,3).value=rs("LINE_NAME")
	sheet.cells(p,4).value=formatdate(rs("START_TIME"),application("longdateformat"))
	sheet.cells(p,5).value=formatdate(rs("CLOSE_TIME"),application("longdateformat"))
	if csng(rs("JOB_START_QUANTITY"))<>0 then
	job_asembly_yield=csng(rs("JOB_GOOD_QUANTITY"))/csng(rs("JOB_START_QUANTITY"))
	else
	job_asembly_yield=0
	end if
	sheet.cells(p,6).value=formatpercent(job_asembly_yield,2,-1)
previous_part_number=rs("PART_NUMBER_ID")
i=p+1
rs.movenext
wend
end if
rs.close

sheet.cells(i,1).value="end"
set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,m-1))
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