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
rs("SHIFT_STATUS")=0
rs.update
end if
rs.close

SQL="select * from LINE_SHIFT"
rs.open SQL,conn,1,3
rs.addnew
rs("LINE_NAME")=line_name
rs("SHIFT_TYPE")="OPEN"
rs("SHIFT_TIME")=now
rs("SHIFT_PERSON")=session("code")
rs.update
rs.close
word="�ɹ����ߡ�"

SQL="select * from JOB where STATUS=5 and LINE_NAME='"&line_name&"'"
rs.open SQL,conn,1,3
if not rs.eof then
i=1
	while not rs.eof
	rs("STATUS")=0
	rs("SHIFT_IN_PERSON")=rs("SHIFT_IN_PERSON")&session("code")&","
	rs("SHIFT_IN_TIME")=rs("SHIFT_IN_TIME")&now&","
	rs.update
	rs.movenext
	i=i+1
	wend
word=word&"�� "&i&" ���������򿪡�"
else
word=word&"��û�й������򿪡�"
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