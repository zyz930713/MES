<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
partnumber=trim(request.Form("partnumber"))
stationname=trim(request.Form("stationname"))
stationseq=trim(request.Form("stationseq"))
traytype=trim(request.Form("traytype"))
traysize=trim(request.Form("traysize"))

SQL="select * from part_tray where part_number='"&partnumber&"' and station_id='"&stationname&"' and tray_type='"&traytype&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	rs("part_number")=partnumber
	rs("station_id")=stationname
	rs("tray_type")=traytype
	rs("station_seq")=stationseq
	rs("tray_size")=traysize
	rs("lm_user")=session("code")
	rs("lm_time")=date()
	rs.update
	rs.close
	word="Successfully edit TrayLink."
	action="location.href='"&beforepath&"'"
else
	word="This TrayLink has not existed, please input again."
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