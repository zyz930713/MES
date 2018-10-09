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
IC_name=request.Form("IC_name")
yearindex=request.Form("yearindex")
weekindex=request.Form("weekindex")
fromdate=request.Form("fromdate")
todate=request.Form("todate")
create_time=now()
NID="IC"&NID_SEQ("IC")

myfile=filePath&"\"&rnd_key&".xls"

SQL="select * from JOB_IC_LIST where TRACKING_NAME='"&IC_name&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("NID")=NID
	rs("REPORT_TIME")=create_time
	rs("TRACKING_NAME")=IC_name
	rs("CREATOR_CODE")=session("code")
	rs("YEAR_INDEX")=yearindex
	rs("WEEK_INDEX")=weekindex
	rs("UPDATE_TIME")=now()
	rs("WEEK_START_DAY")=fromdate
	rs("WEEK_END_DAY")=todate
	rs("EXCEL_FILENAME")=fileName
	rs.update
	rs.close
	
	SQL="delete from JOB_IC_DETAIL where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	SQL="insert into JOB_IC_DETAIL select IC_LOT,IC_QUANTITY,JOB_NUMBER,MACHINE_NUMBER,STATION_START_TIME,IC_PART,PART_NUMBER_ID,CURRENT_STATION_ID,'"&NID&"','"&rnd_key&"','"&session("code")&"' from JOB_IC_TEMP where rnd_key='"&rnd_key&"' and CREATOR_CODE='"&session("code")&"'"
	rs.open SQL,conn,1,3
	
	
	'Create EXCEL file
	set rsU=server.CreateObject("adodb.recordset")
	SQL="select 1,DJI.IC_LOT,DJI.MACHINE_NUMBER,DJI.IC_PART from DISTINCT_JOB_IC_DETAIL DJI where DJI.IC_ID='"&NID&"'"
	rs.Open SQL,conn,1,3
	if not rs.EOF then
	
	filePath=server.mappath("\Reports\Excel")
	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="IC Tracking Report of "&IC_name&" (from "&fromdate&" to "&todate&")"
	sheet.cells(2,1).value="IC Lot Number"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	sheet.cells(2,2).value="IC PART Number"
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
	for n=1 to 10
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
	sheet.cells(i,1).value=rs("IC_LOT")
	sheet.cells(i,2).value=rs("IC_PART")
	sheet.cells(i,3).value=rs("MACHINE_NUMBER")
	
		SQLU="select 1,JI.JOB_NUMBER,JI.STATION_START_TIME,P.PART_NUMBER,JI.CURRENT_STATION_ID,JI.IC_QUANTITY from JOB_IC_DETAIL JI inner join PART P on JI.PART_NUMBER_ID=P.NID where JI.IC_LOT='"&rs("IC_LOT")&"' and JI.IC_ID='"&NID&"' and JI.IC_PART='"&rs("IC_PART")&"' and JI.MACHINE_NUMBER='"&rs("MACHINE_NUMBER")&"' order by JI.JOB_NUMBER"
		rsU.open SQLU,conn,1,3
		if not rsU.eof then
			m=4
			for n=1 to 10
				if not rsU.eof then
				sheet.cells(i,m).value=rsU("JOB_NUMBER")
				m=m+1
				sheet.cells(i,m).value=rsU("PART_NUMBER")
				m=m+1
				sheet.cells(i,m).value=formatdate(rsU("STATION_START_TIME"),application("longdateformat"))
				m=m+1
				sheet.cells(i,m).value=rsU("IC_QUANTITY")
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
	
	savefilePath=server.mappath("\Reports\Tracking\ICTracking\EXCELS")
	myfile=savefilePath&"\"&filename 
	OExcel.ActiveWorkbook.saveas myfile
	OExcel.workbooks.close
	OExcel.quit 
	set sheet=nothing
	set selection=nothing
	set OExcel=Nothing
	
	rs.Close
	set rsU=nothing
	
	word="Successfully save IC report of "&IC_name&" at "&create_time&"."
	action="location.href='"&beforepath&"'"
else
	word="IC report of "&IC_name&" has existed!"
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