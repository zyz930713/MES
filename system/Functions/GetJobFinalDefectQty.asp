<%
function GetJobFinalDefectQty(jobnumber,sheetnumber)	
set rsS=server.CreateObject("adodb.recordset")

SQLS="select * from JOB where JOB_NUMBER='" & jobnumber & "' and SHEET_NUMBER=" & cint(sheetnumber)
rsS.open SQLS,conn,1,3
if not rsS.eof then
	job_routine=rsS("STATIONS_INDEX")
end if
rsS.close

where_stations=""

	if instr(job_routine,"ST00000008")>0 then
		SQLS="select * from STATION where NID='ST00000008'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000008," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if
	
		if instr(job_routine,"ST00000501")>0 then
		SQLS="select * from STATION where NID='ST00000501'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000501," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000043")>0 then
		SQLS="select * from STATION where NID='ST00000043'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000043," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if
	

if instr(job_routine,"ST00000441")>0 then
		SQLS="select * from STATION where NID='ST00000441'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000441," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000385")>0 then
		SQLS="select * from STATION where NID='ST00000385'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000385," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000625")>0 then
		SQLS="select * from STATION where NID='ST00000625'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000625," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
	if instr(job_routine,"ST00000595")>0 then
		SQLS="select * from STATION where NID='ST00000595'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000595," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


if instr(job_routine,"ST00000609")>0 then
		SQLS="select * from STATION where NID='ST00000609'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000609," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000594")>0 then
		SQLS="select * from STATION where NID='ST00000594'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000594," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000597")>0 then
		SQLS="select * from STATION where NID='ST00000597'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000597," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000617")>0 then
		SQLS="select * from STATION where NID='ST00000617'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000617," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if
	          

if instr(job_routine,"ST00000616")>0 then
		SQLS="select * from STATION where NID='ST00000616'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000616," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000322")>0 then
		SQLS="select * from STATION where NID='ST00000322'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000322," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000615")>0 then
		SQLS="select * from STATION where NID='ST00000615'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000615," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	if instr(job_routine,"ST00000618")>0 then
		SQLS="select * from STATION where NID='ST00000618'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000618," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


	if instr(job_routine,"ST00000619")>0 then
		SQLS="select * from STATION where NID='ST00000619'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000619," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


if instr(job_routine,"ST00000597")>0 then
		SQLS="select * from STATION where NID='ST00000597'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000597," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if
	

if instr(job_routine,"ST00000625")>0 then
		SQLS="select * from STATION where NID='ST00000625'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000625," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

	
if instr(job_routine,"ST00000243")>0 then
		SQLS="select * from STATION where NID='ST00000243'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000243," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if



if instr(job_routine,"ST00000627")>0 then
		SQLS="select * from STATION where NID='ST00000627'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000627," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


if instr(job_routine,"ST00000628")>0 then
		SQLS="select * from STATION where NID='ST00000628'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000628," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


if instr(job_routine,"ST00000637")>0 then
		SQLS="select * from STATION where NID='ST00000637'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000637," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000638")>0 then
		SQLS="select * from STATION where NID='ST00000638'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000638," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if


if instr(job_routine,"ST00000640")>0 then
		SQLS="select * from STATION where NID='ST00000640'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000640," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000641")>0 then
		SQLS="select * from STATION where NID='ST00000641'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000641," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if





if instr(job_routine,"ST00000024")>0 then
		SQLS="select * from STATION where NID='ST00000024'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000024," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

if instr(job_routine,"ST00000679")>0 then
		SQLS="select * from STATION where NID='ST00000679'"
		rst.open SQLS,conn,1,3
		if not rst.eof then
			where_stations=where_stations & ",ST00000679," & rst("STATION_ENTER_DEFECTCODE")
		end if
		rst.close
		
	end if

where_stations=replace(mid(where_stations,2),",","','")

'SQLS="Select DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and (STATION_ID in ('ST00000008','ST00000241','ST00000322','ST00000243','ST00000043','ST00000385','ST00000441'))"
SQLS="Select DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&cint(sheetnumber)&" and (STATION_ID in ('" & where_stations & "'))"

rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
	GetJobFinalDefectQty=GetJobFinalDefectQty+cint(rsS("DEFECT_QUANTITY"))
	rsS.movenext
	wend
else
GetJobFinalDefectQty=0
end if
rsS.close
set rsS=nothing

end function
%>