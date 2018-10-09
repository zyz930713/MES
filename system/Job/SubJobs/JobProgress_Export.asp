<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetTempFile.asp" -->
<!--#include virtual="/Components/ExcelConst.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
rnd_key=gen_key(10)
if request("thisdate")<>"" then
thisdate=cdate(request("thisdate"))
else
thisdate=date()
end if
if request("thistime")<>"" then
thistime=cdate(request("thistime"))
else
thistime=#8:00#
end if
if request("hour_select")<>"" then
hour_select=cint(request("hour_select"))
else
hour_select=18
end if
current=hour_select

filePath=server.mappath("\Job\Excel")
set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

set rs1=server.CreateObject("ADODB.RECORDSET")
SQL=session("SQL")
rs.open SQL,conn,1,3

sheet.cells(1,1).value="Job progress on "&thisdate
sheet.cells(2,1).value="No."
sheet.cells(2,1).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,1).Font.ColorIndex = xlColorWhite
sheet.cells(2,1).Font.Bold = True
sheet.cells(2,1).HorizontalAlignment = xlCenter
sheet.cells(2,2).value="Status"
sheet.cells(2,2).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,2).Font.ColorIndex = xlColorWhite
sheet.cells(2,2).Font.Bold = True
sheet.cells(2,2).HorizontalAlignment = xlCenter
sheet.cells(2,3).value="Finished"
sheet.cells(2,3).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,3).Font.ColorIndex = xlColorWhite
sheet.cells(2,3).Font.Bold = True
sheet.cells(2,3).HorizontalAlignment = xlCenter
sheet.cells(2,4).value="Job Number"
sheet.cells(2,4).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,4).Font.ColorIndex = xlColorWhite
sheet.cells(2,4).Font.Bold = True
sheet.cells(2,4).HorizontalAlignment = xlCenter

ktime=thistime
for i=1 to current

sktime=cstr(ktime)
if len(sktime)>7 then
strktime=left(sktime,5)
else
strktime=left(sktime,4)
end if

sheet.cells(2,i+4).value=strktime
sheet.cells(2,i+4).Interior.ColorIndex= xlColorDarkGray
sheet.cells(2,i+4).Font.ColorIndex = xlColorWhite
sheet.cells(2,i+4).Font.Bold = True
sheet.cells(2,i+4).HorizontalAlignment = xlCenter

ktime=dateadd("n",30,ktime)
if ktime>#23:59# then exit for
next

if not rs.eof then
i=3
while not rs.eof 

sheet.cells(i,1).value=i-2

	if rs("STATUS")="0" then
	simg="Opened"
	elseif rs("STATUS")="1" then
	simg="Closed"
	elseif rs("STATUS")="2" then
	simg="Paused"
	elseif rs("STATUS")="3" then
	simg="Locked"
	elseif rs("STATUS")="4" then
	simg="Aborted"
	end if
	  
    sheet.cells(i,2).value=simg
	  
	astations=split(rs("STATIONS_INDEX"),",")
	totalstations=ubound(astations)
	for s=0 to ubound(astations)
		if astations(s)=rs("CURRENT_STATION_ID") then
		cstation=s
		end if
	next
	sheet.cells(i,3).value=formatpercent(cstation/totalstations,,-1)
	
	sheet.cells(i,4).value=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)

	j=0
	id=""
	beforeid=""
	cid=""
	title=""
	beforetitle=""
	ctitle=""
	color=""
	beforecolor=""
	ccol=""
	code=""
	beforecode=""
	ccode=""
	start_time=""
	before_start_time=""
	c_start_time=""
	close_time=""
	before_close_time=""
	c_close_time=""
	cspan=""
	tdout=""
	ktime=thistime
						  
		for s=1 to current
		if ktime>=#23:30# then
		nexttime=#0:0#
		kdate=dateadd("d",1,thisdate)
		else
		nexttime=dateadd("n",30,ktime)
		kdate=thisdate
		end if
			SQL1="select 1,JS.STATION_ID,JS.START_TIME,JS.CLOSE_TIME,JS.OPERATOR_CODE,S.STATION_NAME from JOB_STATIONS JS inner join STATION S on JS.STATION_ID=S.NID where JS.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and JS.SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and ((JS.START_TIME>=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.START_TIME<=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')) or (JS.CLOSE_TIME>=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME<=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')) or (JS.START_TIME<=to_date('"&kdate&" "&ktime&"','yyyy-mm-dd hh24:mi:ss') and JS.CLOSE_TIME>=to_date('"&kdate&" "&nexttime&"','yyyy-mm-dd hh24:mi:ss')))"
			rs1.open SQL1,conn,1,3
			
			if not rs1.eof then
			id=rs1("STATION_ID")
			title=rs1("STATION_NAME")
			code=rs1("OPERATOR_CODE")
			start_time=rs1("START_TIME")
			close_time=rs1("CLOSE_TIME")
				if rs1("STATION_ID")=rs("CURRENT_STATION_ID") then
				color="38"
				else	
				color="35"
				end if
			else
			id=""
			title="&nbsp;"
			code="&nbsp;"
			start_time="&nbsp;"
			close_time="&nbsp;"
			color="0"
			end if
			rs1.close
			
			if id=beforeid and beforeid<>"" then
			j=j+1
			else
				if s<>1 then
				cid=cid&beforeid&","
				ccode=ccode&beforecode&","
				ctitle=ctitle&beforetitle&","
				ccol=ccol&beforecolor&","
				c_start_time=c_start_time&before_start_time&","
				c_close_time=c_close_time&before_close_time&","
				cspan=cspan&j&","
				end if
			j=1
			end if

		beforeid=id
		beforecode=code
		before_start_time=start_time
		before_close_time=close_time
		beforetitle=title
		beforecolor=color
		ktime=dateadd("n",30,ktime)
		lastcurrent=s
		if ktime>#23:59# then exit for  
		next
		
		cid=cid&beforeid&","
		ccode=ccode&beforecode&","
		c_start_time=c_start_time&before_start_time&","
		c_close_time=c_close_time&before_close_time&","
		ctitle=ctitle&beforetitle&","
		ccol=ccol&color&","
		cspan=cspan&j&","
		
		cid=left(cid,len(cid)-1)
		ccode=left(ccode,len(ccode)-1)
		c_start_time=left(c_start_time,len(c_start_time)-1)
		c_close_time=left(c_close_time,len(c_close_time)-1)
		ctitle=left(ctitle,len(ctitle)-1)
		ccol=left(ccol,len(ccol)-1)
		cspan=left(cspan,len(cspan)-1)
		cida=split(cid,",")
		ccodea=split(ccode,",")
		c_start_time_a=split(c_start_time,",")
		c_close_time_a=split(c_close_time,",")
		ctitlea=split(ctitle,",")
		ccola=split(ccol,",")
		cspana=split(cspan,",")
		
		p=0
		q=1

		while q<=lastcurrent
		cell_start=q+4
		cell_end=cell_start+cspana(p)-1
		set selection=sheet.range(sheet.cells(i,cell_start),sheet.cells(i,cell_end))
		selection.merge

		if ccola(p)<>"0" then
		selection.value=ctitlea(p)
		selection.interior.colorIndex=ccola(p)
		'sheet.cells(i,cell_start).addcomment
		notetext=ctitlea(p)&" from "&c_start_time_a(p)&" to "&c_close_time_a(p)&"("&ccodea(p)&")"
    	sheet.range(sheet.cells(i,cell_start),sheet.cells(i,cell_start)).NoteText notetext
		end if
		q=q+cspana(p)
		p=p+1
		wend
		
rs.MoveNext
i=i+1
Wend
rs.Close

sheet.cells(i,1).value="Exported on "&now()
sheet.range(sheet.cells(1,1),sheet.cells(1,lastcurrent+4)).merge
sheet.range(sheet.cells(i,1),sheet.cells(i,lastcurrent+4)).merge
set selection=sheet.range(sheet.cells(1,1),sheet.cells(i,lastcurrent+4))
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

myfile=filePath&"\"&rnd_key&".xls"

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set Sheet=Nothing
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

set rs1=nothing
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->