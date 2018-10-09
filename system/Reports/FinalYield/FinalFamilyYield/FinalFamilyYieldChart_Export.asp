<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!-- #include virtual="/FXChart/Include/ChartFX.ASP.Core.inc" -->
<!-- #include virtual="/FXChart/Include/ChartFX.ASP.Borders.inc" -->
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/getFinalYieldChart.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOLE_Open.asp" -->
<%
factory=Request.Form("factory")
yearnumber=Request.Form("yearnumber")
fromww=cint(Request.Form("fromww"))
toww=cint(Request.Form("toww"))
imagewidth=cint(Request.Form("imagewidth"))
imageheight=cint(Request.Form("imageheight"))
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")
If Request.Form("exportfamily") = "1" Then
	seriesgroup="OVERALL"&","&getSeriesGroup("TEXT","","","",",")
	seriesgroup=left(seriesgroup,len(seriesgroup)-1)
	a_seriesgroup=split(seriesgroup,",")
	for s=0 to ubound(a_seriesgroup)-1
		SQL = "Select FINAL_YIELD,FAMILY_TARGET_YIELD,INPUT_QUANTITY,OUTPUT_QUANTITY,CHART_WEEK,CHART_YEAR,CHART_MONTH from FAMILYYIELD_WEEK where FACTORY_ID='"&factory&"' and CHART_YEAR="&yearnumber&" and CHART_WEEK is not null and FAMILY_NAME='"&a_seriesgroup(s)&"' and CHART_WEEK>="&fromww& " and CHART_WEEK<="&toww&" order by CHART_WEEK"
		getFinalYieldChart a_seriesgroup(s),SQL,imagewidth,imageheight,filePath,rnd_key&repeatstring(s,"0",3)&".Jpeg"
	next
else
	seriesgroup=Request.Form("seriesgroup")
	SQL = "Select FINAL_YIELD,FAMILY_TARGET_YIELD,INPUT_QUANTITY,OUTPUT_QUANTITY,CHART_WEEK,CHART_YEAR,CHART_MONTH from FAMILYYIELD_WEEK where FACTORY_ID='"&factory&"' and CHART_YEAR="&yearnumber&" and CHART_WEEK is not null and FAMILY_NAME='"&seriesgroup&"' and CHART_WEEK>="&fromww&" and CHART_WEEK<="&toww&" order by CHART_WEEK"
	getFinalYieldChart seriesgroup,SQL,imagewidth,imageheight,filePath,"this.Jpeg"
end if
'generate excel
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet
myLeft = 10
myTop = 10
myWidth = imagewidth
myHeight = imageheight

If Request.Form("exportfamily") = "1" Then
	For s=0 To ubound(a_seriesgroup)-1
	Sheet.Shapes.AddPicture filePath&"\"&rnd_key&repeatstring(s,"0",3)&".Jpeg", False, True, myLeft, myTop, myWidth, myHeight
	myTop=myTop+myHeight+10
	Next
else
	Sheet.Shapes.AddPicture filePath&"\this.Jpeg", False, True, myLeft, myTop, myWidth, myHeight
end if

myfile=filePath&"\"&rnd_key&".xls"
OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=Nothing
set selection=nothing
set OExcel=Nothing

'Create a stream object
Set Stream = Server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile myfile
bytes=Stream.Read
Stream.Close
Set Stream = Nothing

DeleteFile filePath&"\*.xls"
DeleteFile filePath&"\*.Jpeg" '删除该目录下所有原先产生的临时打印文件 

'Output the contents of the stream object
Response.ContentType = "application/vnd.ms-excel"
Response.BinaryWrite bytes
response.end
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->