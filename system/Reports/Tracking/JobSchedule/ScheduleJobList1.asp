<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Tracking/JobSchedule/ScheduleCheck.asp" -->
<!--#include virtual="/Job/IsAdmin.asp" -->
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
words=""
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		SQL="select PART_NUMBER_ID from JOB_SCHEDULE_DETAIL where JOB_NUMBER='"&request.Form("jobnumber"&i)&"'"
		rs.open SQL,conn,1,3
		rs("PART_NUMBER_ID")=request.Form("part_number_id"&i)
		rs.update
		rs.close
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%end if%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%
response.Redirect(beforepath)%>