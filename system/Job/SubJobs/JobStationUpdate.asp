<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
'update start quantity and defect quantity for each station after this station
SQL="select STATIONS_INDEX,REPEATED_STATIONS_SEQUENCE from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
stations_index=rs("STATIONS_INDEX")
repeated_sequence=rs("REPEATED_STATIONS_SEQUENCE")
end if
rs.close

a_stations_index=split(stations_index,",")
a_repeated_sequence=split(repeated_sequence,",")

total_defect_quantity=0
this_defect_quantity=0
for i=0 to ubound(a_stations_index)
	if i=0 then
	SQL="select ACTION_VALUE from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&a_repeated_sequence(i)&" and ACTION_ID='AC00000097'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	first_action_start_quantity=cint(rs("ACTION_VALUE"))
	end if
	rs.close
	SQL="select STATION_START_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&a_repeated_sequence(i)
	rs.open SQL,conn,1,3
	if not rs.eof then
		first_station_start_quantity=cint(rs("STATION_START_QUANTITY"))
		if first_station_start_quantity<>first_action_start_quantity then
		first_station_start_quantity=first_action_start_quantity
		end if
		next_start_quantity=first_station_start_quantity
	end if
	rs.close
	end if
	SQL="select nvl(sum(DEFECT_QUANTITY),0) as DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE='"&a_repeated_sequence(i)&"'"
	rs.open SQL,conn,1,3
	this_defect_quantity=cint(rs("DEFECT_QUANTITY"))
	rs.close
	this_start_quantity=next_start_quantity
	this_good_quantity=this_start_quantity-this_defect_quantity
	total_defect_quantity=total_defect_quantity+this_defect_quantity
	next_start_quantity=this_good_quantity
	SQL="Update JOB_ACTIONS set ACTION_VALUE='"&this_start_quantity&"' where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&a_repeated_sequence(i)&" and ACTION_ID='AC00000097'"
	rs.open SQL,conn,1,3
	SQL="select * from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"' and STATION_ID='"&a_stations_index(i)&"' and REPEATED_SEQUENCE="&a_repeated_sequence(i)
	rs.open SQL,conn,1,3
	if not rs.eof then
	rs("STATION_START_QUANTITY")=this_start_quantity
	rs("GOOD_QUANTITY")=this_good_quantity
	rs("STATION_DEFECTCODE_QUANTITY")=this_good_quantity
	rs.update
	end if
	rs.close
next

SQL="select * from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
session("aerror")=SQL
rs.open SQL,conn,1,3
if not rs.eof then
old_total_defect_quantity=cint(rs("JOB_DEFECTCODE_QUANTITY"))
diff=total_defect_quantity-old_total_defect_quantity
rs("JOB_START_QUANTITY")=first_station_start_quantity
rs("JOB_GOOD_QUANTITY")=first_station_start_quantity-total_defect_quantity
rs("JOB_DEFECTCODE_QUANTITY")=total_defect_quantity
rs("JOB_ASSEMBLY_YIELD")=(first_station_start_quantity-total_defect_quantity)/first_station_start_quantity
rs.update
end if
rs.close
SQL="update JOB_MASTER set DEFECTCODE_QUANTITY=DEFECTCODE_QUANTITY+"&diff&" where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>