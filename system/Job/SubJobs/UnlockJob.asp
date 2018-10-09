<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
reason=request.QueryString("reason")
SQL="select STATUS,CURRENT_STATION_ID,INTERVAL_SKIP_STATION_ID from JOB where JOB_NUMBER='"&jobnumber&"' and sheetnumber='"&sheetnumber&"'"
rs.open SQL,conn,1,3
current_station_id=rs("CURRENT_STATION_ID")
rs("STATUS")=0
rs("INTERVAL_SKIP_STATION_ID")=rs("CURRENT_STATION_ID")
rs.update
rs.close
SQL="select CONTROL_TYPE,CONTROL_STATION,CONTROL_REASON,CONTROL_PERSON,CONTROL_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and sheetnumber='"&sheetnumber&"'"
rs.open SQL,conn,1,3
rs("CONTROL_TYPE")=rs("CONTROL_TYPE")&"3"&","
rs("CONTROL_STATION")=rs("CONTROL_STATION")&current_station_id&","
rs("CONTROL_REASON")=rs("CONTROL_REASON")&reason&","
rs("CONTROL_PERSON")=rs("CONTROL_PERSON")&session("code")&","
rs("CONTROL_TIME")=rs("CONTROL_TIME")&now()&","
rs.update
rs.close
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%response.Redirect(beforepath)%>