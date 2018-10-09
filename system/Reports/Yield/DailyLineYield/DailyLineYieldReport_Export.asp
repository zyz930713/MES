<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
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

family_names=getSeriesGroup("TEXT",""," where S.FACTORY_ID='"&factory_id&"'"," order by S.SERIES_GROUP_NAME","|")
if family_names<>"" then
family_names=left(family_names,len(family_names)-1)
a_family_names=split(family_names,"|")
end if
line_names=getLine("TEXT",""," where L.FACTORY_ID='"&factory_id&"'"," order by L.LINE_NAME",",")
if line_names<>"" then
line_names=left(line_names,len(line_names)-1)
a_line_names=split(line_names,",")
end if
SQL="select TARGET_FIRSTYIELD from FACTORY where NID='"&factory_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_first_target=csng(rs("TARGET_FIRSTYIELD"))/100
else
factory_first_target=0
end if
rs.close

	set OExcel=createobject("Excel.application")
	OExcel.Workbooks.open(filePath&"\book1.xlt")
	OExcel.sheets(1).select
	set Sheet=OExcel.activeWorkbook.ActiveSheet
	
	sheet.cells(1,1).value="Daily Line Yield Report from "&formatdate(from_time,application("longdateformat"))&" to "&formatdate(to_time,application("longdateformat"))
	
	sheet.cells(2,1).value="Family Name"
	sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,1).Font.ColorIndex = xlColorWhite
	sheet.cells(2,1).Font.Bold = True
	sheet.cells(2,1).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,1),sheet.cells(3,1)).merge
	sheet.cells(2,2).value="Line Name"
	sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,2).Font.ColorIndex = xlColorWhite
	sheet.cells(2,2).Font.Bold = True
	sheet.cells(2,2).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,2),sheet.cells(3,2)).merge
	
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=3
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
	sheet.range(sheet.cells(2,i),sheet.cells(2,i+4)).merge
	current_date=dateadd("d",1,current_date)
	i=i+5
	loop
	sheet.cells(2,i).value="Overall"
	sheet.cells(2,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(2,i).Font.ColorIndex = xlColorWhite
	sheet.cells(2,i).Font.Bold = True
	sheet.cells(2,i).HorizontalAlignment = xlCenter
	sheet.range(sheet.cells(2,i),sheet.cells(2,i+4)).merge
	
	current_date=cdate(from_time)
	end_date=cdate(dateadd("d",-1,to_time))
	i=3
	do while current_date<=end_date
	sheet.cells(3,i).value="Input"
	sheet.cells(3,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i).Font.Bold = True
	sheet.cells(3,i).HorizontalAlignment = xlCenter
	sheet.cells(3,i+1).value="Output"
	sheet.cells(3,i+1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+1).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+1).Font.Bold = True
	sheet.cells(3,i+1).HorizontalAlignment = xlCenter
	sheet.cells(3,i+2).value="Yield"
	sheet.cells(3,i+2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+2).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+2).Font.Bold = True
	sheet.cells(3,i+2).HorizontalAlignment = xlCenter
	sheet.cells(3,i+3).value="Target"
	sheet.cells(3,i+3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+3).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+3).Font.Bold = True
	sheet.cells(3,i+3).HorizontalAlignment = xlCenter
	sheet.cells(3,i+4).value="Gap"
	sheet.cells(3,i+4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+4).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+4).Font.Bold = True
	sheet.cells(3,i+4).HorizontalAlignment = xlCenter
	current_date=dateadd("d",1,current_date)
	i=i+5
	loop
	sheet.cells(3,i).value="Input"
	sheet.cells(3,i).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i).Font.Bold = True
	sheet.cells(3,i).HorizontalAlignment = xlCenter
	sheet.cells(3,i+1).value="Output"
	sheet.cells(3,i+1).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+1).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+1).Font.Bold = True
	sheet.cells(3,i+1).HorizontalAlignment = xlCenter
	sheet.cells(3,i+2).value="Yield"
	sheet.cells(3,i+2).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+2).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+2).Font.Bold = True
	sheet.cells(3,i+2).HorizontalAlignment = xlCenter
	sheet.cells(3,i+3).value="Target"
	sheet.cells(3,i+3).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+3).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+3).Font.Bold = True
	sheet.cells(3,i+3).HorizontalAlignment = xlCenter
	sheet.cells(3,i+4).value="Gap"
	sheet.cells(3,i+4).Interior.ColorIndex= xlColorDarkGray
	sheet.cells(3,i+4).Font.ColorIndex = xlColorWhite
	sheet.cells(3,i+4).Font.Bold = True
	sheet.cells(3,i+4).HorizontalAlignment = xlCenter

m=4
overall_input_quantity=0
overall_output_quanityt=0
overall_yield=0
set rs1=server.CreateObject("adodb.recordset")
for i=0 to ubound(a_family_names)
	for j=0 to ubound(a_line_names)
	SQL="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and LINE_NAME='"&a_line_names(j)&"' and LINE_DAY>=to_date('"&from_time&"','yyyy-mm-dd hh24:mi:ss') and LINE_DAY<=to_date('"&to_time&"','yyyy-mm-dd hh24:mi:ss')"
	rs.open SQL,conn,1,3
	if not rs.eof then
		line_input_quantity=0
		line_output_quantity=0
		line_yield=0
		line_target_yield=0
		sheet.cells(m,1).value=a_family_names(i)
		sheet.cells(m,2).value=a_line_names(j)
		current_date=cdate(from_time)
		end_date=cdate(dateadd("d",-1,to_time))
		k=3
		do while current_date<=end_date
	SQL1="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='"&a_family_names(i)&"' and LINE_NAME='"&a_line_names(j)&"' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
		rs1.open SQL1,conn,1,3
		if not rs1.eof then
		assembly_input_quantity=csng(rs1("ASSEMBLY_INPUT_QUANTITY"))
		assembly_output_quantity=csng(rs1("ASSEMBLY_OUTPUT_QUANTITY"))
		assembly_yield=csng(rs1("ASSEMBLY_YIELD"))
		family_target_firstyield=csng(rs1("FAMILY_TARGET_FIRSTYIELD"))
		else
		assembly_input_quantity=0
		assembly_output_quantity=0
		assembly_yield=0
		family_target_firstyield=0
		end if
		rs1.close
		sheet.cells(m,k).value=assembly_input_quantity
		sheet.cells(m,k+1).value=assembly_output_quantity
		sheet.cells(m,k+2).value=formatpercent(assembly_yield,2,-1)
		sheet.cells(m,k+2).Interior.ColorIndex=getNormalYieldExcelColor(assembly_yield,family_target_firstyield)
		sheet.cells(m,k+3).value=formatpercent(family_target_firstyield,2,-1)
		sheet.cells(m,k+4).value=formatpercent((assembly_yield-family_target_firstyield),2,-1)
		line_input_quantity=line_input_quantity+assembly_input_quantity
		line_output_quantity=line_output_quantity+assembly_output_quantity
		overall_input_quantity=overall_input_quantity+assembly_input_quantity
		overall_output_quantity=overall_output_quantity+assembly_output_quantity
		if family_target_firstyield<>0 then
		line_target_yield=family_target_firstyield
		end if
		current_date=dateadd("d",1,current_date)
		k=k+5
		loop
		if line_input_quantity<>0 then
		line_yield=line_output_quantity/line_input_quantity
		end if
		sheet.cells(m,k).value=line_input_quantity
		sheet.cells(m,k+1).value=line_output_quantity
		sheet.cells(m,k+2).value=formatpercent(line_yield,2,-1)
		sheet.cells(m,k+2).Interior.ColorIndex=getNormalYieldExcelColor(line_yield,line_target_yield)
		sheet.cells(m,k+3).value=formatpercent(line_target_yield,2,-1)
		sheet.cells(m,k+4).value=formatpercent((line_yield-line_target_yield),2,-1)
	m=m+1
	end if
	rs.close
	next
next

'overall yield
sheet.cells(m,1).value="Overall"
sheet.cells(m,1).Interior.ColorIndex=xlColorRose
sheet.cells(m,2).Interior.ColorIndex=xlColorRose
current_date=cdate(from_time)
end_date=cdate(dateadd("d",-1,to_time))
i=3
do while current_date<=end_date
	SQL1="select * from DAILY_FAMILYLINE_YIELD where PROFILE_TASK_ID='"&profile_task_id&"' and FACTORY_ID='"&factory_id&"' and FAMILY_NAME='OVERALL' and LINE_NAME='OVERALL' and to_char(LINE_DAY,'yyyy-mm-dd')='"&formatdate(current_date,application("F_shortdateformat"))&"'"
rs1.open SQL1,conn,1,3
if not rs1.eof then
total_input_quantity=csng(rs1("ASSEMBLY_INPUT_QUANTITY"))
total_output_quantity=csng(rs1("ASSEMBLY_OUTPUT_QUANTITY"))
total_yield=csng(rs1("ASSEMBLY_YIELD"))
else
total_input_quantity=0
total_output_quantity=0
total_yield=0
end if
rs1.close
sheet.cells(m,i).value=total_input_quantity
sheet.cells(m,i).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+1).value=total_output_quantity
sheet.cells(m,i+1).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+2).value=formatpercent(total_yield,2,-1)
sheet.cells(m,i+2).Interior.ColorIndex=getNormalYieldExcelColor(total_yield,factory_first_target)
sheet.cells(m,i+3).value=formatpercent(factory_first_target,2,-1)
sheet.cells(m,i+3).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+4).value=formatpercent((total_yield-factory_first_target),2,-1)
sheet.cells(m,i+4).Interior.ColorIndex=xlColorRose
current_date=dateadd("d",1,current_date)
i=i+5
loop
if overall_input_quantity<>0 then
overall_yield=overall_output_quantity/overall_input_quantity
end if
sheet.cells(m,i).value=overall_input_quantity
sheet.cells(m,i).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+1).value=overall_output_quantity
sheet.cells(m,i+1).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+2).value=formatpercent(overall_yield,2,-1)
sheet.cells(m,i+2).Interior.ColorIndex=getNormalYieldExcelColor(overall_yield,factory_first_target)
sheet.cells(m,i+3).value=formatpercent(factory_first_target,2,-1)
sheet.cells(m,i+3).Interior.ColorIndex=xlColorRose
sheet.cells(m,i+4).value=formatpercent((overall_yield-factory_first_target),2,-1)
sheet.cells(m,i+4).Interior.ColorIndex=xlColorRose

set rs1=nothing
	'export time
	sheet.cells(m+1,1).value="Exported on "&now()
	sheet.range(sheet.cells(1,1),sheet.cells(1,i+4)).merge
	sheet.range(sheet.cells(m+1,1),sheet.cells(m+1,i+4)).merge
	set selection=sheet.range(sheet.cells(1,1),sheet.cells(m+1,i+4))
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