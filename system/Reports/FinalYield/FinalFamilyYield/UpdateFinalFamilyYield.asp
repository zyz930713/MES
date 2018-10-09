<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath="FinalFamilyYieldList.asp"
finalfamily_id=request.Form("finalfamily_id")
SQL="select * from FINAL_FAMILYYIELD_LIST where NID='"&finalfamily_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	if request.Form("isww")="1" then
	rs("CHART_WEEK")=request.Form("wwnumber")
	rs("CHART_YEAR")=request.Form("yearnumber")
	rs("CHART_MONTH")=request.Form("monthnumber")
	else
	rs("CHART_WEEK")=null
	rs("CHART_YEAR")=null
	rs("CHART_MONTH")=null
	end if
	rs.update
word="Successfully Update FINAL FAMILY YIELD."
action="window.close()"
else
word="This FINAL FAMILY YIELD is not existed."
action="window.close()"
end if
rs.close
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
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