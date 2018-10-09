<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Yield Exclusion/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
yield_exclusion_name=trim(request.Form("yield_exclusion_name"))
SQL="select * from YIELD_EXCLUSION where YIELD_EXCLUSION_NAME='"&yield_exclusion_name&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="YE"&NID_SEQ("YIELD_EXCLUSION")
rs("YIELD_EXCLUSION_NAME")=yield_exclusion_name
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
rs("INCLUDED_SYSTEM_ITEMS")=replace(request.Form("toitem")," ","")
rs.update
rs.close
word="Successfully copy this Yield Exclusion."
action="location.href='"&beforepath&"'"
else
word="Yield Exclusion of "&yield_exclusion_name&" has existed, please input again."
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