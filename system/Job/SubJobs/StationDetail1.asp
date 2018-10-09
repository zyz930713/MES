<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->

<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
if request.Form("actionidcount")<>"" or request.Form("defectidcount")<>"" then
%>
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
jobnumber=request.Form("jobnumber")
sheetnumber=request.Form("sheetnumber")
station_id=request.Form("station_id")
jobtype=request.Form("jobtype")
repeated_sequence=request.Form("repeated_sequence")
for i=1 to request.Form("actionidcount")
	SQL="select * from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&station_id&"' and REPEATED_SEQUENCE="&repeated_sequence&" and ACTION_ID='"&request.Form("action_id"&i)&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	rs("ACTION_VALUE")=request.Form("action_value"&i)
	rs.update
	else
	rs.addnew
	rs("JOB_NUMBER")=jobnumber
	rs("SHEET_NUMBER")=sheetnumber
	rs("JOB_TYPE")=jobtype
	rs("STATION_ID")=station_id
	rs("REPEATED_SEQUENCE")=repeated_sequence
	rs("ACTION_ID")=request.Form("action_id"&i)
	rs("ACTION_VALUE")=request.Form("action_value"&i)
	rs.update
	end if
	rs.close
	'Get Mother Action ID
		SQL="select action_purpose from action where nid='"+request.Form("action_id"&i)+"'"
		rs.open SQL,conn,1,3
		actPurpose=""
		if rs.recordcount<>0 then
			actPurpose=rs("action_purpose")
		end if
		rs.close
	'end

	if actPurpose="5" then
		SQL="Update JOB_STATIONS set GOOD_QUANTITY='"&request.Form("action_value"&i)&"' where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&station_id&"' and REPEATED_SEQUENCE="&repeated_sequence
		rs.open SQL,conn,1,3
	end if
next
for i=1 to request.Form("defectidcount")
	deftQty=request.Form("defectcode_value"&i)
	SQL="select * from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and STATION_ID='"&station_id&"' and REPEATED_SEQUENCE="&repeated_sequence&" and DEFECT_CODE_ID='"&request.Form("defectcode_id"&i)&"'"
	rs.open SQL,conn,1,3
	if rs.eof and deftQty <>"" then
		rs.addnew
		rs("JOB_NUMBER")=jobnumber
		rs("SHEET_NUMBER")=sheetnumber
		rs("REPEATED_SEQUENCE")=repeated_sequence
		rs("STATION_ID")=station_id
		rs("DEFECT_CODE_ID")=request.Form("defectcode_id"&i)
	end if
	if not rs.eof and deftQty="" then
		deftQty="-"&rs("DEFECT_QUANTITY")
		rs.delete
	elseif deftQty <>"" then
		deftQty=csng(cint(deftQty)-cint(rs("DEFECT_QUANTITY")))
		rs("DEFECT_QUANTITY")=request.Form("defectcode_value"&i)		
		rs.update
	end if		
	rs.close
	
	if deftQty <>"" then
		SQL="select job_number,batchno,qty,subjoblist,actualqty,operator,remark,intimestamp,transactiontype,isprint from offline_store "
		SQL=SQL+"where isprint=0 and remark='"&request.Form("defectcode_id"&i)&"' and job_number='"&jobnumber&"' and subjoblist='"&sheetnumber&"' "
		rs.open SQL,conn,1,3
		if rs.eof then
			rs.addnew
			rs("job_number")=jobnumber
			rs("batchno")=jobnumber
			rs("qty")=deftQty
			rs("subjoblist")=sheetnumber
			rs("actualqty")=deftQty
			rs("operator")=session("code")
			rs("remark")=request("defectcode_id"&i)
			rs("intimestamp")=now()
			rs("transactiontype")=request("trans_type"&i)
			rs("isprint")="0"
		else		
			rs("qty")=cint(rs("qty")) + cint(deftQty)
			rs("actualqty")=cint(rs("actualqty"))+cint(deftQty)
		end if
		if rs("qty")="0" then
			rs.delete
		else
			rs.update
		end if
		rs.close
	end if
next
SystemLog "Job - Defectcode","Change defectcode of "&jobnumber&"-"&repeatstring(sheetnumber,"0",3)&" ("&station_id&")"
'delete temp defect
SQL="delete from job_defectcodes_temp where job_number='"&jobnumber&"' and sheet_number="&sheetnumber&" and station_id='"&station_id&"'"
rs.open SQL,conn,1,3

'update start quantity and defect quantity for each station after this station
SQL="select STATIONS_INDEX from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
stations_index=rs("STATIONS_INDEX")
end if
rs.close

a_stations_index=split(stations_index,",")
current_station=0
total_defect_quantity=0
for i=0 to ubound(a_stations_index)
	if station_id=a_stations_index(i) then
	current_station=i
	end if
	if i>=current_station then
		if i=0 then
			this_defect_quantity=0			
			SQL="select ACTION_VALUE from JOB_ACTIONS a ,action b where a.action_id=b.nid(+) and a.JOB_NUMBER='"&jobnumber&"' and a.SHEET_NUMBER="&sheetnumber&" and a.JOB_TYPE='"&jobtype&"' and a.STATION_ID='"&a_stations_index(i)&"' and a.REPEATED_SEQUENCE="&repeated_sequence&" and b.action_purpose=5"
			rs.open SQL,conn,1,3
			if not rs.eof then
			first_action_start_quantity=clng(rs("ACTION_VALUE"))
			end if
			rs.close
			SQL="select STATION_START_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&repeated_sequence
			rs.open SQL,conn,1,3
			if not rs.eof then
				first_station_start_quantity=clng(rs("STATION_START_QUANTITY"))
				if first_station_start_quantity<>first_action_start_quantity then
				first_station_start_quantity=first_action_start_quantity
				end if
				next_start_quantity=first_station_start_quantity
			else
			rs.close
			exit for
			end if
			rs.close
		end if
		SQL="select nvl(sum(DEFECT_QUANTITY),0) as DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE='"&repeated_sequence&"'"
		rs.open SQL,conn,1,3
		this_defect_quantity=clng(rs("DEFECT_QUANTITY"))
		rs.close
		this_start_quantity=next_start_quantity
		this_good_quantity=this_start_quantity-this_defect_quantity
		total_defect_quantity=total_defect_quantity+this_defect_quantity
		next_start_quantity=this_good_quantity
		SQL="Update JOB_ACTIONS set ACTION_VALUE='"&this_start_quantity&"' where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&repeated_sequence&" AND action_id in ( select nid from action where action_purpose=5)"

		rs.open SQL,conn,1,3
		SQL="select * from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&repeated_sequence
		rs.open SQL,conn,1,3
		if not rs.eof then
		rs("STATION_START_QUANTITY")=this_start_quantity
		rs("GOOD_QUANTITY")=this_good_quantity
		rs("STATION_DEFECTCODE_QUANTITY")=this_defect_quantity
		rs.update
		else
		rs.close
		exit for
		end if
		rs.close
	else
		SQL="select GOOD_QUANTITY,STATION_DEFECTCODE_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&repeated_sequence
		rs.open SQL,conn,1,3
		if not rs.eof then
		next_start_quantity=clng(rs("GOOD_QUANTITY"))
		total_defect_quantity=total_defect_quantity+clng(rs("GOOD_QUANTITY"))
		else
		rs.close
		exit for
		end if
		rs.close
	end if
next

LineName=""

SQL="select * from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
session("aerror")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
old_total_defect_quantity=clng(rs("JOB_DEFECTCODE_QUANTITY"))
diff=total_defect_quantity-old_total_defect_quantity
rs("JOB_START_QUANTITY")=first_station_start_quantity
rs("JOB_GOOD_QUANTITY")=first_station_start_quantity-total_defect_quantity
rs("JOB_DEFECTCODE_QUANTITY")=total_defect_quantity
rs("JOB_ASSEMBLY_YIELD")=(first_station_start_quantity-total_defect_quantity)/first_station_start_quantity
Line_Name=rs("Line_Name")

rs.update
end if
rs.close
SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("DEFECTCODE_QUANTITY")=csng(rs("DEFECTCODE_QUANTITY"))+diff
rs("ASSEMBLY_GOOD_QUANTITY")=csng(rs("ASSEMBLY_GOOD_QUANTITY"))+diff
	if csng(rs("ASSEMBLY_INPUT_QUANTITY"))<>0 then
	rs("ASSEMBLY_YIELD")=(csng(rs("ASSEMBLY_GOOD_QUANTITY"))+diff)/csng(rs("ASSEMBLY_INPUT_QUANTITY"))
	end if
rs.update
end if
rs.close


%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->

 
<%end if%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>