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
MEMS_name=request.Form("MEMS_name")
yearindex=request.Form("yearindex")
weekindex=request.Form("weekindex")
fromdate=request.Form("fromdate")
todate=request.Form("todate")
create_time=now()
NID="ME"&NID_SEQ("MEMS")

myfile=filePath&"\"&rnd_key&".xls" 

SQL="select * from JOB_MEMS_LIST where TRACKING_NAME='"&MEMS_name&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("NID")=NID
	rs("REPORT_TIME")=create_time
	rs("TRACKING_NAME")=MEMS_name
	rs("CREATOR_CODE")=session("code")
	rs("YEAR_INDEX")=yearindex
	rs("WEEK_INDEX")=weekindex
	rs("UPDATE_TIME")=now()
	rs("WEEK_START_DAY")=fromdate
	rs("WEEK_END_DAY")=todate
	rs("EXCEL_FILENAME")=fileName
	rs.update
	rs.close
	
	SQL="delete from JOB_MEMS_DETAIL where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	SQL="insert into JOB_MEMS_DETAIL select MEMS_LOT,MEMS_QUANTITY,JOB_NUMBER,MACHINE_NUMBER,STATION_START_TIME,MEMS_PART,PART_NUMBER_ID,CURRENT_STATION_ID,'"&NID&"','"&rnd_key&"','"&session("code")&"' from JOB_MEMS_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	
	
	'Create EXCEL file
	set rsU=server.CreateObject("adodb.recordset")
	SQL="select 1,DJM.MEMS_LOT,DJM.MACHINE_NUMBER,DJM.MEMS_PART from DISTINCT_JOB_MEMS_DETAIL DJM where DJM.MEMS_ID='"&NID&"'"
	rs.Open SQL,conn,1,3
	if not rs.EOF then
	
	filePath=server.mappath("\Reports\Excel")
	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="MEMS Tracking Report of "&MEMS_name&" (from "&fromdate&" to "&todate&")"
	sheet.cells(2,1).value="MEMS Lot Number"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	sheet.cells(2,2).value="MEMS PART Number"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	sheet.cells(2,3).value="Machine NO"
	sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,3).Font.ColorIndex = xlColorWhite
	sheet.cells(2,3).Font.Bold = True
	sheet.cells(2,3).HorizontalAlignment = xlCenter
	m=4
	for n=1 to 5
	sheet.cells(2,m).value="Job "&n
	sheet.cells(2,m).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,m).Font.ColorIndex = xlColorWhite
	sheet.cells(2,m).Font.Bold = True
	sheet.cells(2,m).HorizontalAlignment = xlCenter
	m=m+1
	sheet.cells(2,m).value="Part "&n
	sheet.cells(2,m).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,m).Font.ColorIndex = xlColorWhite
	sheet.cells(2,m).Font.Bold = True
	sheet.cells(2,m).HorizontalAlignment = xlCenter
	m=m+1
	sheet.cells(2,m).value="Station Time "&n
	sheet.cells(2,m).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,m).Font.ColorIndex = xlColorWhite
	sheet.cells(2,m).Font.Bold = True
	sheet.cells(2,m).HorizontalAlignment = xlCenter
	m=m+1
	sheet.cells(2,m).value="Quantity "&n
	sheet.cells(2,m).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,m).Font.ColorIndex = xlColorWhite
	sheet.cells(2,m).Font.Bold = True
	sheet.cells(2,m).HorizontalAlignment = xlCenter
	m=m+1
	next
	lastm=m
	
	i=3
	while Not rs.EOF
	sheet.cells(i,1).value=rs("MEMS_LOT")
	sheet.cells(i,2).value=rs("MEMS_PART")
	sheet.cells(i,3).value=rs("MACHINE_NUMBER")
	
		SQLU="select 1,JM.JOB_NUMBER,JM.STATION_START_TIME,P.PART_NUMBER,JM.CURRENT_STATION_ID,JM.MEMS_QUANTITY from JOB_MEMS_DETAIL JM inner join PART P on JM.PART_NUMBER_ID=P.NID where JM.MEMS_LOT='"&rs("MEMS_LOT")&"' and JM.MEMS_ID='"&NID&"' and JM.MEMS_PART='"&rs("MEMS_PART")&"' and JM.MACHINE_NUMBER='"&rs("MACHINE_NUMBER")&"' order by JM.JOB_NUMBER"
		rsU.open SQLU,conn,1,3
		if not rsU.eof then
			m=4
			for n=1 to 5
				if not rsU.eof then
				sheet.cells(i,m).value=rsU("JOB_NUMBER")
				m=m+1
				sheet.cells(i,m).value=rsU("PART_NUMBER")
				m=m+1
				sheet.cells(i,m).value=formatdate(rsU("STATION_START_TIME"),application("longdateformat"))
				m=m+1
				sheet.cells(i,m).value=rsU("MEMS_QUANTITY")
				m=m+1
				rsU.movenext
				end if
			next
		end if
		rsU.close
	i=i+1
	rs.MoveNext
	wend
	
	sheet.cells(i,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,lastm-1)).merge
	sheet.range(sheet.cells(i,1),sheet.cells(i,lastm-1)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,lastm-1))
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
	
	savefilePath=server.mappath("\Reports\Tracking\MEMSTracking\EXCELS")
	myfile=savefilePath&"\"&filename 
	OExcel.ActiveWorkbook.saveas myfile
	OExcel.workbooks.close
	OExcel.quit 
	set sheet=nothing
	set selection=nothing
	set OExcel=Nothing
	
	rs.Close
	set rsU=nothing
	
	word="Successfully save MEMS report of "&MEMS_name&" at "&create_time&"."
	action="location.href='"&beforepath&"'"
else
	word="MEMS report of "&MEMS_name&" has existed!"
	action="location.href='"&beforepath&"'"
rs.close
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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