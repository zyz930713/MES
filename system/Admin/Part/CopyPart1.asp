<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
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
rs("STATIONS_COUNT")=request.Form("stationscount")
stations_index=replace(request.Form("toitem")," ","")
rs("STATIONS_INDEX")=stations_index
rs("STATIONS_ROUTINE")=request.Form("stationroutine")
rs("STATIONS_TRANSACTION")=getStationsTransaction(stations_index)
rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
rs("TRACKING_CODE")=trim(request.Form("trackingcode"))
rs("PART_PER_BOARD")=trim(request.Form("part_per_board"))
rs("PART_PER_FRAME")=trim(request.Form("part_per_frame"))
rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
rs.update
rs.close
SQL="insert * into STATION from STATION where NID in ('"&replace(toitem,",","','")&"')"
rs.open SQL,conn,1,3
word="Successfully copy this Part."
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
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->