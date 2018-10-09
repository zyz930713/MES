<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Scrap/ScrapCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
trans=false
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
	trans=true
	exit for
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
end if
words=""
j=1
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		SQL="update JOB_MASTER set START_ERROR=0 where JOB_NUMBER='"&request.Form("jobnumber"&i)&"'"
		rs.open SQL,conn,1,3
		j=j+1
	end if
next
word="Successfully cancel Start error of "&j-1&"."
action="location.href='"&beforepath&"'"
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%end if%>
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