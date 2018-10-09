<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
upcount=0
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
partnumber=request.Form("partnumber")
part_rule=trim(request.Form("part_rule"))
SQL="select * from PART where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	rs("PART_NUMBER")=partnumber
	rs("ROUTINE_PURPOSE")=request.Form("routine_purpose")
	rs("STATIONS_COUNT")=request.Form("stationscount")
	lines_index=replace(request.Form("toitem1")," ","")
	rs("LINES_INDEX")=lines_index
	stations_index=replace(request.Form("toitem2")," ","")
	old_stations_index=rs("STATIONS_INDEX")
	rs("STATIONS_INDEX")=stations_index
	rs("STATIONS_ROUTINE")=request.Form("stationroutine")
	rs("STATIONS_TRANSACTION")=getStationsTransaction(stations_index)
	rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
	if trim(request.Form("initial_quantity"))<>"" then
	rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
	else
	rs("INITIAL_QUANTITY")=null
	end if
	rs("TARGET_YIELD")=request.Form("yield")
	rs("MEET_PRIORITY")=request.Form("priority")
	rs.update
	word="Successfully edit Part."
	action="location.href='"&beforepath&"'"
else
	word="Part of "&partnumber&" has not existed, please input again."
	action="history.back()"
end if
rs.close%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%
if old_stations_index<>stations_index then
	set rsc=server.CreateObject("adodb.recordset")
	SQLC="select 1,count(*) as upcount from JOB where PART_NUMBER_ID='"&id&"' and STATUS=0"
	rsc.open SQLC,conn,1,3
	if not rsc.eof then
	upcount=rsc("upcount")
	else
	upcount=0
	end if
	rsc.close
	set rsc=nothing
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%if upcount<>0 then%>
if (confirm("There are <%=upcount%> opened jobs applied to old routine! Would you update them?"))
{location.href='EditPartUpdateJob.asp?id=<%=id%>&path=<%=path%>&query=<%=query%>'}
else
{<%=action%>;}
<%else%>
<%=action%>;
<%end if%>
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->