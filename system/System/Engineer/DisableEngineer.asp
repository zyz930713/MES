<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SystemLog.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
rolename=request.QueryString("rolename")
SQL="update users set STATUS=0 where NID='"&id&"'"
rs.open SQL,conn,3,3
SystemLog "System - Engineer","Disabled Engineer of "&rolename&" ("&id&")"
word="Engineer of "&rolename&" is disabled!"
%>
<script language="javascript">
alert("<%=word%>");
location.href='<%=beforepath%>';
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->