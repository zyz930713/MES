<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.Form("jobnumber")
sheetnumber=request.Form("sheetnumber")
station_to_change=replace(request.Form("station_id")," ","")
SQL="select STATUS,CURRENT_STATION_ID,LAST_STATION_ID,FINISHED_STATIONS_ID,STATIONS_INDEX,STATIONS_TRANSACTION from JOB where JOB_NUMBER='"&jobnumber&"' and sheetnumber='"&sheetnumber&"'"
rs.open SQL,conn,1,3
rs("STATUS")=0
astation_to_change=split(station_to_change,",")
astation=split(rs("STATIONS_INDEX"),",")
afinished_station_id=split(left(rs("FINISHED_STATIONS_ID"),len(rs("FINISHED_STATIONS_ID"))-1),",")
atransaction=split(rs("STATIONS_TRANSACTION"),",")
rs("CURRENT_STATION_ID")=astation_to_change(0)
set rsS=server.CreateObject("adodb.recordset")
for i=0 to ubound(astation) 'update transaction index
	for j=0 to ubound(astation_to_change)
		if astation(i)=astation_to_change(j) then
			atransaction(i)="0"
			SQLS="delete JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and sheetnumber='"&sheetnumber&"' and STATION_ID='"&astation_to_change(j)&"'"
			rsS.open SQLS,conn,1,3
'			SQLS="delete JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and STATION_ID='"&astation_to_change(j)&"'"
'			rsS.open SQLS,conn,1,3
		end if
	next
	transaction=transaction&atransaction(i)&","
next
set rsS=nothing
transaction=left(transaction,len(transaction)-1)
rs("STATIONS_TRANSACTION")=transaction

for i=0 to ubound(afinished_station_id) 'update finished stations index
	if afinished_station_id(i)=astation_to_change(0) then
	exit for
	else
	finished=finished&afinished_station_id(i)&","
	end if
next
rs("FINISHED_STATIONS_ID")=finished

	for i=ubound(astation) to lbound(astation) step -1 'update the last station
		if atransaction(i)="0" then
		rs("LAST_STATION_ID")=astation(i)
		exit for
		end if
	next
rs.update
rs.close
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>