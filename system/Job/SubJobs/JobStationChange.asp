<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
station_id=request.QueryString("station_id")
change_type=request.QueryString("type")
SQL="select LAST_STATION_ID,STATIONS_INDEX,STATIONS_TRANSACTION from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"'"
rs.open SQL,conn,1,3
astation=split(rs("STATIONS_INDEX"),",")
atransaction=split(rs("STATIONS_TRANSACTION"),",")
for i=0 to ubound(astation)
	if astation(i)=station_id then
		if change_type="0" then
		atransaction(i)="0"
		else
		atransaction(i)="1"
		end if
	end if
	transaction=transaction&atransaction(i)&","
next
transaction=left(transaction,len(transaction)-1)
rs("STATIONS_TRANSACTION")=transaction
	for i=ubound(astation) to lbound(astation) step -1
		if atransaction(i)="0" then
		rs("LAST_STATION_ID")=astation(i)
		exit for
		end if
	next
rs.update
rs.close
SystemLog "Job - Job","Change station("&station_id&")'s transaction of "&jobnumber&"-"&sheetnumber
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>