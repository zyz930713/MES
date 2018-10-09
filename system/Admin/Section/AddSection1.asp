<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Section/SectionCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
sectionname=trim(request.Form("sectionname"))
SQL="select * from SECTION where SECTION_NAME='"&sectionname&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="SE"&NID_SEQ("SECTION")
rs("SECTION_NAME")=sectionname
rs("FACTORY_ID")=request.Form("factory")
rs.update
rs.close
word="Successfully save a New Section."
action="location.href='"&beforepath&"'"
else
word="Section of "&sectionname&" has existed, please input again."
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