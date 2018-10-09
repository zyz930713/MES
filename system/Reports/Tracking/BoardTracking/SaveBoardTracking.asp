<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobTracking/JobTrackingCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetTempFile.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
rnd_key=request.Form("rnd_key")
BOARD_name=request.Form("BOARD_name")
yearindex=request.Form("yearindex")
weekindex=request.Form("weekindex")
fromdate=request.Form("fromdate")
todate=request.Form("todate")
create_time=now()
NID="BO"&NID_SEQ("BOARD")

myfile=filePath&"\"&rnd_key&".xls"

SQL="select * from JOB_BOARD_LIST where TRACKING_NAME='"&BOARD_name&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("NID")=NID
	rs("REPORT_TIME")=create_time
	rs("TRACKING_NAME")=BOARD_name
	rs("CREATOR_CODE")=session("code")
	rs("YEAR_INDEX")=yearindex
	rs("WEEK_INDEX")=weekindex
	rs("UPDATE_TIME")=now()
	rs("WEEK_START_DAY")=fromdate
	rs("WEEK_END_DAY")=todate
	rs("EXCEL_FILENAME")=fileName
	rs.update
	rs.close
	
	SQL="delete from JOB_BOARD_DETAIL where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	SQL="insert into JOB_BOARD_DETAIL select BOARD_LOT,BOARD_PART,JOB_NUMBER,START_TIME,PART_NUMBER_ID,CURRENT_STATION_ID,'"&NID&"','"&rnd_key&"','"&session("code")&"' from JOB_BOARD_RAW where START_TIME>=TO_DATE('"&fromdate&"','YYYY-MM-DD') and START_TIME<=TO_DATE('"&todate&"','YYYY-MM-DD')"
	rs.open SQL,conn,1,3
	
	
	'Create EXCEL file
	set rsU=server.CreateObject("adodb.recordset")
	SQL="select 1,JB.*,P.PART_NUMBER from JOB_BOARD_DETAIL JB inner join PART P on JB.PART_NUMBER_ID=P.NID where JB.BOARD_ID='"&NID&"'"
	rs.Open SQL,conn,1,3
	if not rs.EOF then
	
	filePath=server.mappath("\Reports\Excel")
	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Board Tracking Report of "&BOARD_name&" (from "&fromdate&" to "&todate&")"
	sheet.cells(2,1).value="Job Number"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	sheet.cells(2,2).value="Part Number"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	sheet.cells(2,3).value="Board Lot Number"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	sheet.cells(2,4).value="Board PART Number"
	sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,4).Font.ColorIndex = xlColorWhite
	sheet.cells(2,4).Font.Bold = True
	sheet.cells(2,4).HorizontalAlignment = xlCenter
	sheet.cells(2,5).value="Start Time"
	sheet.cells(2,5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,5).Font.ColorIndex = xlColorWhite
	sheet.cells(2,5).Font.Bold = True
	sheet.cells(2,5).HorizontalAlignment = xlCenter
	
	i=3
	while Not rs.EOF
	sheet.cells(i,1).value=rs("JOB_NUMBER")
	sheet.cells(i,2).value=rs("PART_NUMBER")
	sheet.cells(i,3).value=rs("BOARD_LOT")
	sheet.cells(i,4).value=rs("BOARD_PART")
	sheet.cells(i,5).value=formatdate(rs("START_TIME"),application("longdateformat"))
	i=i+1
	rs.MoveNext
	wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,5)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,5)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,5))
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
	
	savefilePath=server.mappath("\Reports\Tracking\BoardTracking\EXCELS")
	myfile=savefilePath&"\"&filename 
	OExcel.ActiveWorkbook.saveas myfile
	OExcel.workbooks.close
	OExcel.quit 
	set sheet=nothing
	set selection=nothing
	set OExcel=Nothing
	
	rs.Close
	set rsU=nothing
	
	word="Successfully save Board report of "&BOARD_name&" at "&create_time&"."
	action="location.href='"&beforepath&"'"
else
	word="BOARD report of "&BOARD_name&" has existed!"
	action="location.href='"&beforepath&"'"
rs.close
end if
%>
<!DOCTYPE HTML PUBLBOARD "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->