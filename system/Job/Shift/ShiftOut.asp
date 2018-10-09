<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
pagename="ShiftIn.asp"
id=request.QueryString("id")
line_name=request.QueryString("line_name")

SQL="select * from LINE where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("SHIFT_STATUS")=1
rs.update
end if
rs.close

SQL="select * from LINE_SHIFT"
rs.open SQL,conn,1,3
rs.addnew
rs("LINE_NAME")=line_name
rs("SHIFT_TYPE")="STOP"
rs("SHIFT_TIME")=now
rs("SHIFT_PERSON")=session("code")
rs.update
rs.close
word="成功关线。"

SQL="select * from JOB where STATUS=0 and LINE_NAME='"&line_name&"'"
rs.open SQL,conn,1,3
if not rs.eof then
i=1
	while not rs.eof
	rs("STATUS")=5
	rs("SHIFT_OUT_PERSON")=rs("SHIFT_OUT_PERSON")&session("code")&","
	rs("SHIFT_OUT_TIME")=rs("SHIFT_OUT_TIME")&now&","
	rs.update
	rs.movenext
	i=i+1
	wend
word=word&"有 "&i&" 个工单被暂时关闭。"
else
word=word&"但没有工单被暂时关闭。"
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>");
location.href="<%=beforepath%>";
</script>
</head>

<body>
</body>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->