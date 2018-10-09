<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Reports/Yield/FinanceCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/Functions/GetYieldColor.asp" -->
<%
rnd_key=gen_key(10)
filePath=server.mappath("\Reports\Excel")

path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
from_time=request.QueryString("from_time")
to_time=request.QueryString("to_time")
profile_task_id=trim(request.QueryString("profile_task_id"))
factory_id=trim(request.QueryString("factory_id"))

family_names=getFinanceSeriesGroup("TEXT",""," where S.FACTORY_ID='"&factory_id&"'"," order by S.SERIES_GROUP_NAME","|")
if family_names<>"" then
family_names=left(family_names,len(family_names)-1)
a_family_names=split(family_names,"|")
end if
SQL="select FINANCE_TARGET_YIELD,FINANCE_PLAN_TARGET_YIELD from FACTORY where NID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_finance_target=csng(rs("FINANCE_TARGET_YIELD"))
factory_finance_plan_target=csng(rs("FINANCE_PLAN_TARGET_YIELD"))
else
factory_finance_target=0
factory_finance_plan_target=0
end if
rs.close

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Daily Finance Yield Report from "&formatdate(from_time,application("longdateformat"))&" to "&formatdate(to_time,application("longdateformat"))
	
	sheet.cells(2,1).value="Family Name"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,1),sheet.cells(3,1)).merge
	
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=2
	do while current_date<=end_date
		if weekday(current_date)=1 then
		datacolor=xlColorRed
		elseif weekday(current_date)=7 then
		datacolor=xlColorGreen
		else
		datacolor=xlColorWhite
		end if
	sheet.cells(2,i).value=current_date&" ("&shortweekdayconvert(weekday(current_date))&")"
	sheet.cells(2,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,i).Font.ColorIndex = datacolor
	sheet.cells(2,i).Font.Bold = True
	sheet.cells(2,i).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,i),sheet.cells(2,i+9)).merge
	current_date=dateadd("d",1,current_date)
	i=i+6
	loop
	sheet.cells(2,i).value="Overall"
	sheet.cells(2,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,i).Font.ColorIndex = xlColorWhite
	sheet.cells(2,i).Font.Bold = True
	sheet.cells(2,i).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,i),sheet.cells(2,i+9)).merge
	
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=2
	do while current_date<=end_date
	sheet.cells(3,i).value="Input Value"
	sheet.cells(3,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i).Font.Bold = True
	sheet.cells(3,i).HorizontalAlignment = xlCenter
	sheet.cells(3,i+1).value="Output Value"
	sheet.cells(3,i+1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+1).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+1).Font.Bold = True
	sheet.cells(3,i+1).HorizontalAlignment = xlCenter
	sheet.cells(3,i+2).value="Scrap Value"
	sheet.cells(3,i+2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+2).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+2).Font.Bold = True
	sheet.cells(3,i+2).HorizontalAlignment = xlCenter
	sheet.cells(3,i+3).value="Yield"
	sheet.cells(3,i+3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+3).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+3).Font.Bold = True
	sheet.cells(3,i+3).HorizontalAlignment = xlCenter
	sheet.cells(3,i+4).value="Target"
	sheet.cells(3,i+4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+4).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+4).Font.Bold = True
	sheet.cells(3,i+4).HorizontalAlignment = xlCenter
	sheet.cells(3,i+5).value="Plan Target"
	sheet.cells(3,i+5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+5).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+5).Font.Bold = True
	sheet.cells(3,i+5).HorizontalAlignment = xlCenter
	sheet.cells(3,i+6).value="Gap"
	sheet.cells(3,i+6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+6).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+6).Font.Bold = True
	sheet.cells(3,i+6).HorizontalAlignment = xlCenter
	current_date=dateadd("d",1,current_date)
	i=i+7
	loop
	sheet.cells(3,i).value="Input Value"
	sheet.cells(3,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i).Font.Bold = True
	sheet.cells(3,i).HorizontalAlignment = xlCenter
	sheet.cells(3,i+1).value="Output Value"
	sheet.cells(3,i+1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+1).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+1).Font.Bold = True
	sheet.cells(3,i+1).HorizontalAlignment = xlCenter
	sheet.cells(3,i+2).value="Scrap Value"
	sheet.cells(3,i+2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+2).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+2).Font.Bold = True
	sheet.cells(3,i+2).HorizontalAlignment = xlCenter
	sheet.cells(3,i+3).value="Yield"
	sheet.cells(3,i+3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+3).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+3).Font.Bold = True
	sheet.cells(3,i+3).HorizontalAlignment = xlCenter
	sheet.cells(3,i+4).value="Target"
	sheet.cells(3,i+4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+4).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+4).Font.Bold = True
	sheet.cells(3,i+4).HorizontalAlignment = xlCenter
	sheet.cells(3,i+5).value="Plan Target"
	sheet.cells(3,i+5).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+5).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+5).Font.Bold = True
	sheet.cells(3,i+5).HorizontalAlignment = xlCenter
	sheet.cells(3,i+6).value="Gap"
	sheet.cells(3,i+6).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+6).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+6).Font.Bold = True
	sheet.cells(3,i+6).HorizontalAlignment = xlCenter

m=4
overall_input_quantity=0
overall_output_quanityt=0
overall_yield=0
for i=0 to ubound(a_family_names)
	line_input_quantity=0
	line_output_quantity=0
	line_yield=0
	line_target_yield=0
	sheet.cells(m,1).value=a_family_names(i)
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	k=2
	do while current_date<=end_date
	SQL="select * from DAILY_FINANCE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	input_value=ccur(rs("INPUT_AMOUNT"))
	output_value=ccur(rs("OUTPUT_AMOUNT"))
	scrap_value=ccur(rs("SCRAP_AMOUNT"))
	amount_yield=csng(rs("AMOUNT_YIELD"))
	family_target_yield=csng(rs("TARGET_YIELD"))
	family_plan_target_yield=csng(rs("PLAN_TARGET_YIELD"))
	else
	stock_input_quantity=0
	stock_output_quantity=0
	stock_yield=0
	family_internal_yield=0
	family_target_yield=0
	family_plan_target_yield=0
	end if
	rs.close
	sheet.cells(m,k).value=cstr(input_value)
	sheet.cells(m,k+1).value=cstr(output_value)
	sheet.cells(m,k+2).value=cstr(scrap_value)
	sheet.cells(m,k+3).value=formatpercent(amount_yield,2,-1)
	sheet.cells(m,k+3).Interior.ColorIndex=getNormalYieldExcelColor(amount_yield,family_target_yield)
	sheet.cells(m,k+4).value=formatpercent(family_target_yield,2,-1)
	sheet.cells(m,k+5).value=formatpercent(family_plan_target_yield,2,-1)
	sheet.cells(m,k+6).value=formatpercent((amount_yield-family_target_yield),2,-1)
	line_input_value=line_input_value+input_value
	line_output_value=line_output_value+output_value
	line_scrap_value=line_scrap_value+scrap_value
	overall_input_value=overall_input_value+input_value
	overall_output_value=overall_output_value+output_value
	overall_scrap_value=overall_scrap_value+scrap_value
	if family_target_yield<>0 then
	line_target_yield=family_target_yield
	end if
	if family_plan_target_yield<>0 then
	line_plan_target_yield=family_plan_target_yield
	end if
	current_date=dateadd("d",1,current_date)
	k=k+7
	loop
	if line_input_value<>0 then
	line_yield=line_output_value/line_input_value
	end if
	sheet.cells(m,k).value=cstr(line_intput_value)
	sheet.cells(m,k+1).value=cstr(line_output_value)
	sheet.cells(m,k+2).value=cstr(line_scrap_value)
	sheet.cells(m,k+3).value=formatpercent(line_yield,2,-1)
	sheet.cells(m,k+3).Interior.ColorIndex=getNormalYieldExcelColor(line_yield,line_target_yield)
	sheet.cells(m,k+4).value=formatpercent(line_target_yield,2,-1)
	sheet.cells(m,k+5).value=formatpercent(line_plan_target_yield,2,-1)
	sheet.cells(m,k+6).value=formatpercent((line_yield-line_target_yield),2,-1)
	m=m+1
next

'overall yield
sheet.cells(m,1).value="Overall"
sheet.cells(m,1).Interior.ColorIndex=xlColorRose
current_date=cdate(from_time)
end_date=cdate(dateadd("d",-1,to_time))
i=2
do while current_date<=end_date
	SQL="select * from DAILY_FINANCE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='OVERALL' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
if not rs.eof then
total_input_value=ccur(rs("INPUT_AMOUNT"))
total_output_value=csng(rs("OUTPUT_AMOUNT"))
total_scrap_value=csng(rs("SCRAP_AMOUNT"))
total_yield=csng(rs("AMOUNT_YIELD"))
else
total_input_value=0
total_output_value=0
total_scrap_value=0
total_yield=0
end if
rs.close
sheet.cells(m,i).value=cstr(total_input_value)
sheet.cells(m,i).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+1).value=cstr(total_output_value)
sheet.cells(m,i+1).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+2).value=cstr(total_scrap_value)
sheet.cells(m,i+2).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+3).value=formatpercent(total_yield,2,-1)
sheet.cells(m,i+3).Interior.ColorIndex=getNormalYieldExcelColor(total_yield,factory_target)
sheet.cells(m,i+3).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+4).value=formatpercent(factory_target,2,-1)
sheet.cells(m,i+4).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+5).value=formatpercent(factory_plan_target,2,-1)
sheet.cells(m,i+5).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+6).value=formatpercent((total_yield-factory_target),2,-1)
sheet.cells(m,i+6).Interior.ColorIndex=xlColorRose
current_date=dateadd("d",1,current_date)
i=i+7
loop
if overall_input_amount<>0 then
overall_yield=overall_output_amount/overall_input_amount
end if
sheet.cells(m,i).value=cstr(overall_input_value)
sheet.cells(m,i).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+1).value=cstr(overall_output_value)
sheet.cells(m,i+1).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+2).value=cstr(overall_scrap_value)
sheet.cells(m,i+2).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+3).value=formatpercent(overall_yield,2,-1)
sheet.cells(m,i+3).Interior.ColorIndex=getNormalYieldExcelColor(overall_yield,factory_target)
sheet.cells(m,i+3).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+4).value=formatpercent(factory_target,2,-1)
sheet.cells(m,i+4).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+5).value=formatpercent(factory_target,2,-1)
sheet.cells(m,i+5).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+6).value=formatpercent((overall_yield-family_target_firstyield),2,-1)
sheet.cells(m,i+6).Interior.ColorIndex=xlColorRose

set rs=nothing
	'export time
	sheet.cells(m+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,i+9)).merge
	sheet.range(sheet.cells(m+1,1),sheet.cells(m+1,i+9)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(m+1,i+9))
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
	
	DeleteFile filePath&"\*.xls" '删除该目录下所有原先产生的临时打印文件 
	
	'Output the contents of the stream object
	Response.ContentType = "application/vnd.ms-excel"
	Response.BinaryWrite bytes
	response.end
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->