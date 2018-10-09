<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=trim(request.QueryString("jobnumber"))
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
if DBA=true then
	SQL="select CLOSE_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		if request.QueryString("close_time")<>"" then
		rs("CLOSE_TIME")=request.QueryString("close_time")
		else
		rs("CLOSE_TIME")=null
		end if
	rs.update
	end if
	rs.close
end if
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>