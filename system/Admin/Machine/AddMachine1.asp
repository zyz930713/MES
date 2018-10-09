<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Machine/MachineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
machinenumber=request.Form("machinenumber")
SQL="select * from MACHINE where MACHINE_NUMBER='"&machinenumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="MA"&NID_SEQ("MACHINE")
rs("MACHINE_NUMBER")=machinenumber
rs("MACHINE_NAME")=request.Form("machinename")
rs("FACTORY_ID")=request.Form("factory")
rs("STATIONS_USED")=replace(request.Form("toitem")," ","")
	if request.Form("locked")="1" then
	rs("LOCKED")=1
	else
	rs("LOCKED")=0
	end if
rs.update
rs.close
word="Successfully save a New Machine."
action="location.href='"&beforepath&"'"
else
word="Machine of "&machinenumber&" has existed, please input again."
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