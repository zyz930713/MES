<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
if DBA=true then
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
'backup data
SQL="insert into JOB_BACKUP select * from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER="&sheetnumber&" and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="insert into JOB_STATIONS_BACKUP select * from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="insert into JOB_ACTIONS_BACKUP select * from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="insert into JOB_ACTIONS_REPEATED_BACKUP select * from JOB_ACTIONS_REPEATED where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="insert into JOB_DEFECTCODES_BACKUP select * from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"'"
rs.open SQL,conn,1,3

'delete data
SQL="delete from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="delete from JOB_STATIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="delete from JOB_ACTIONS where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="delete from JOB_ACTIONS_REPEATED where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
SQL="delete from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"'"
rs.open SQL,conn,1,3
'log
SystemLog "Job - Job","Delete Job of "&jobnumber&"-"&replace(jobtype,"N","")&repeatstring(sheetnumber,"0",3)
end if
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>