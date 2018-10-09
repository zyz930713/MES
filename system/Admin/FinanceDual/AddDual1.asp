<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
dual=trim(request.Form("dual"))
SQL="select * from DUAL_SETTINGS where DUAL_NAME='"&dual&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="DU"&NID_SEQ("DUAL_SETTINGS")
rs("DUAL_NAME")=dual
rs("SINGLE1")=request.Form("Single1")
rs("SINGLE2")=request.Form("Single2")
rs.update
rs.close
word="Successfully save a Dual Settings."
action="location.href='"&beforepath&"'"
else
word="Dual Settings of "&dual&" has existed, please input again."
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