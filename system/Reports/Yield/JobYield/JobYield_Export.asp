<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library"-->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
server.ScriptTimeout=999999999
filePath=server.mappath("\Reports\Excel")
SQL=session("SQL")
set rsD=server.CreateObject("adodb.recordset")
set rst=server.CreateObject("adodb.recordset")
set rsSD=server.CreateObject("adodb.recordset")
'FactoryRight "D."

set OExcel=createobject("Excel.application")
OExcel.Workbooks.open(filePath&"\book1.xlt")
OExcel.sheets(1).select
set Sheet=OExcel.activeWorkbook.ActiveSheet

rs.Open SQL,conn,1,3
main_job_qty=0
job_qty=0


last_job=""
	

do while not rs.eof
		
	job_qty=job_qty+1
	
	if rs("JOB_NUMBER")<>current_job then
		main_job_qty=main_job_qty+1

		job_routine=rs("STATIONS_INDEX")
		
		where_stations=""
		
		if instr(job_routine,"ST00000008")>0 then
			SQLS="select * from STATION where NID='ST00000008'"
			rst.open SQLS,conn,1,3
			if not rst.eof then
				where_stations=where_stations & "," & rst("STATION_ENTER_DEFECTCODE")
			end if
			rst.close
			
		end if
		
		if instr(job_routine,"ST00000043")>0 then
			SQLS="select * from STATION where NID='ST00000043'"
			rst.open SQLS,conn,1,3
			if not rst.eof then
				where_stations=where_stations & "," & rst("STATION_ENTER_DEFECTCODE")
			end if
			rst.close
			
		end if
		
		if instr(job_routine,"ST00000441")>0 then
			SQLS="select * from STATION where NID='ST00000441'"
			rst.open SQLS,conn,1,3
			if not rst.eof then
				where_stations=where_stations & "," & rst("STATION_ENTER_DEFECTCODE")
			end if
			rst.close
			
		end if
		
		if instr(job_routine,"ST00000385")>0 then
			SQLS="select * from STATION where NID='ST00000385'"
			rst.open SQLS,conn,1,3
			if not rst.eof then
				where_stations=where_stations & "," & rst("STATION_ENTER_DEFECTCODE")
			end if
			rst.close
			
		end if
		
		where_stations=replace(mid(where_stations,2),",","','")
		
		SQLSD="select sum(JOB_START_QUANTITY) as total_QTY from JOB where JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
		rst.open SQLSD,conn,1,3
		total_QTY=rst("total_QTY")
		rst.close
		
		SQL="select * from DEFECTCODE D where (STATION_ID in ('" & where_stations & "'))"
	
		rst.open SQL,conn,1,3
			
		sheet.cells(job_qty,1).value="Job Number"
		sheet.cells(job_qty,2).value="Part Number"
		sheet.cells(job_qty,3).value="Start Time"
		sheet.cells(job_qty,4).value="Start Quantity"
		sheet.range(sheet.cells(job_qty,5),sheet.cells(job_qty,rst.recordcount+4)).merge
		sheet.cells(job_qty,5).value="Total Defect Code"
		sheet.cells(job_qty,rst.recordcount+5).value="Yield"
		sheet.cells(job_qty,rst.recordcount+6).value="Solderer"
		sheet.cells(job_qty+1,1).value=rs("Job_Number")
		sheet.cells(job_qty+1,2).value=rs("PART_NUMBER_TAG")
		sheet.cells(job_qty+1,3).value=rs("Start_Time")

		sheet.cells(job_qty+1,4).value=total_QTY

		df_qty=0
		do while not rst.eof
			
			sheet.cells(job_qty+1,5+df_qty).value=rst("DEFECT_CHINESE_NAME")
			
			df_qty=df_qty+1
			rst.movenext
		loop
		rst.close
		
		SQLSD="select nvl(sum(JD.defect_quantity),0) as total_dfs from JOB_DEFECTCODES JD where JD.JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
		rst.open SQLSD,conn,1,3
		
		sheet.cells(job_qty+1,df_qty+5).value=formatpercent((clng(total_QTY)-clng(rst("total_dfs")))/clng(total_QTY),2)
		
		rst.close
		
		SQL="select * from JOB_STATIONS where JOB_NUMBER='" & rs("Job_Number") & "' and SHEET_NUMBER=" & clng(rs("SHEET_NUMBER")) & " and STATION_ID in ('ST00000085','ST00000696','ST00000141','ST00000201','ST00000302','ST00000381','ST00000421')"
		rst.open SQL,conn,1,3
		if not rst.eof then
		sheet.cells(job_qty+1,df_qty+6).value=rsT("OPERATOR_CODE")
		end if
		rst.close
	else
	
		SQL="select * from DEFECTCODE D where (STATION_ID in ('" & where_stations & "'))"
	
		rst.open SQL,conn,1,3
				
		sheet.cells(job_qty+1,1).value=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
		sheet.cells(job_qty+1,2).value=rs("PART_NUMBER_TAG")
		sheet.cells(job_qty+1,3).value=formatdate(rs("START_TIME"),application("longdateformat"))
		sheet.cells(job_qty+1,4).value=rs("JOB_START_QUANTITY")
		df_qty=0
		total_dfs=0
		do while not rst.eof
			SQLSD="select * from JOB_DEFECTCODES JD where JD.JOB_NUMBER='"&rs("JOB_NUMBER")&"' and JD.SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and JD.DEFECT_CODE_ID='" & trim(rst("NID")&"") & "'"
			rsSD.open SQLSD,conn,1,3
			if not rsSD.eof then
			DEFECT_QUANTITY=rsSD("DEFECT_QUANTITY")
			sheet.cells(job_qty+1,5+df_qty).value=rsSD("DEFECT_QUANTITY")
			else
			DEFECT_QUANTITY=0
			sheet.cells(job_qty+1,5+df_qty).value=0
			end if
			df_qty=df_qty+1
			total_dfs=clng(total_dfs)+clng(DEFECT_QUANTITY)
			rsSD.close
			rst.movenext
		loop		

		'sheet.cells(job_qty+1,rst.recordcount+5).value=formatpercent(rs("JOB_ASSEMBLY_YIELD"),1)
		sheet.cells(job_qty+1,rst.recordcount+5).value=formatpercent((clng(rs("JOB_START_QUANTITY"))-clng(total_dfs))/clng(rs("JOB_START_QUANTITY")),2)
		rst.close
		SQL="select * from JOB_STATIONS where JOB_NUMBER='" & rs("Job_Number") & "' and SHEET_NUMBER=" & clng(rs("SHEET_NUMBER")) & " and STATION_ID in ('ST00000085','ST00000696','ST00000141','ST00000201','ST00000302','ST00000381','ST00000421')"
		rst.open SQL,conn,1,3
		if not rst.eof then
		sheet.cells(job_qty+1,df_qty+6).value=rsT("OPERATOR_CODE")
		end if
		rst.close
		
		last_job=rs("JOB_NUMBER")
		rs.movenext	
	end if
	if not rs.eof then
		current_job=rs("JOB_NUMBER")
	end if
loop



rs.Close


set myFs=server.createObject("scripting.FileSystemObject") 
fileName=getTempFile(myFs) '取得一个临时文件名 
myfile=filePath&filename 

OExcel.ActiveWorkbook.saveas myfile
OExcel.workbooks.close
OExcel.quit 
set sheet=nothing
set OExcel=Nothing

function getTempFile(myFileSystem)  
dim tempFile,dotPos 
tempFile=myFileSystem.getTempName 
dotPos=instr(1,tempFile,".") 
getTempFile=mid(tempFile,1,dotPos)&"xls" 
end function

'Create a stream object
Set Stream = Server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile myfile
bytes=Stream.Read
Stream.Close
Set Stream = Nothing

set myFs=server.createObject("scripting.FileSystemObject") 
myFs.DeleteFile filePath&"*.xls" '删除该目录下所有原先产生的临时打印文件 
set myFs=nothing 

'Output the contents of the stream object
Response.ContentType = "application/vnd.ms-excel"
Response.BinaryWrite bytes
response.end

set rsD=nothing
set rsSD=nothing
set rst=nothing
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->