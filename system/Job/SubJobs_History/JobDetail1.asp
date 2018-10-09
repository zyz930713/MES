<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
trans=false
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
	trans=true
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
end if
jobnumber=request.Form("jobnumber")
sheetnumber=request.Form("sheetnumber")
jobtype=request.Form("jobtype")
 
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		SQL="select JOB_NUMBER,STATION_ID,OPERATOR_CODE,STATUS,START_TIME,CLOSE_TIME from bar_hist.JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"' and STATION_ID='"&request.Form("station_id"&i)&"' and REPEATED_SEQUENCE='"&request.Form("repeated_sequence"&i)&"'"
		 
		rs.open SQL,conn,1,3
		if not rs.eof then
		rs("OPERATOR_CODE")=request.Form("operator_code"&i)
		rs("START_TIME")=request.Form("start_time"&i)
		rs("CLOSE_TIME")=request.Form("stop_time"&i)
		rs.update
		else
		rs.addnew
		rs("JOB_NUMBER")=jobnumber
		rs("STATION_ID")=request.Form("station_id"&i)
		rs("STATUS")=2
		rs("OPERATOR_CODE")=request.Form("operator_code"&i)
		rs("START_TIME")=request.Form("start_time"&i)
		rs("CLOSE_TIME")=request.Form("stop_time"&i)
		rs.update
		end if
		rs.close
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%end if%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>