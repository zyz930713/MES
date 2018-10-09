<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
eventname=trim(request.Form("name"))
SQL="select * from EVENT where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
rs("EVENT_NAME")=eventname
rs("DESCRIPTION")=trim(request.Form("description"))
note=trim(request.Form("note"))
note=replace(note," ","&nbsp;")
note=replace(note, vbCrLf, "<BR>")
rs("EVENT_NOTE")=note
body=trim(request.Form("body"))
body=replace(body," ","&nbsp;")
body=replace(body, vbCrLf, "<BR>")
rs("EMAIL_BODY")=body
rs("EMAIL_HYPERLINK")=trim(request.Form("hyperlink"))
rs.update
word="Successfully edit Event."
action="location.href='"&beforepath&"'"
else
word="Event of "&eventname&" has existed, please input again."
action="history.back()"
end if
rs.close
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="JavaScript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->