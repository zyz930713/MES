<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
partnumber=trim(request.Form("partnumber"))
SQL="select * from PART where PART_NUMBER='"&partnumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="PA"&NID_SEQ("PART")
rs("PART_NUMBER")=partnumber
rs("PART_RULE")=request.Form("part_rule")
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
rs("ROUTINE_TYPE")=1
rs("ROUTINE_PURPOSE")=request.Form("routine_purpose")
rs("STATIONS_COUNT")=request.Form("stationscount")
lines_index=replace(request.Form("toitem1")," ","")
rs("LINES_INDEX")=lines_index
stations_index=replace(request.Form("toitem2")," ","")
rs("STATIONS_INDEX")=stations_index
rs("STATIONS_ROUTINE")=request.Form("stationroutine")
rs("STATIONS_TRANSACTION")=getStationsTransaction(stations_index)
rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
rs("TRACKING_CODE")=trim(request.Form("trackingcode"))
if request.Form("part_per_board")<>"" then
rs("PART_PER_BOARD")=trim(request.Form("part_per_board"))
end if
if request.Form("part_per_frame")<>"" then
rs("PART_PER_FRAME")=trim(request.Form("part_per_frame"))
end if
if request.Form("initial_quantity")<>"" then
rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
end if
rs("TARGET_YIELD")=request.Form("yield")
rs("MEET_PRIORITY")=request.Form("priority")
rs.update
rs.close
word="Successfully save a New Part."
action="location.href='"&beforepath&"'"
else
word="Part of "&partnumber&" has existed, please input again."
action="history.back()"
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
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->