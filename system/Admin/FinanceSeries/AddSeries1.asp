<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/FinanceSeries/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
seriesname=trim(request.Form("seriesname"))
SQL="select * from FINANCE_SERIES where SERIES_NAME='"&seriesname&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
NID="FS"&NID_SEQ("FINANCE_SERIES")
rs("NID")=NID
rs("SERIES_NAME")=seriesname
rs("FACTORY_ID")=request.Form("factory")
rs("TARGET_YIELD")=csng(request.Form("yield"))/100
rs.update
rs.close
SQL="update PRODUCT_MODEL set SERIES_ID='"&id&"' where instr('"&request.Form("toitem")&"',ITEM_ID)>0"
rs.open SQL,conn,1,3
SQL="update USERS set SERIES_FILTER='"&request.Form("models_filter_notin")&"' where USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
word="Successfully save a New Finance Series."
action="location.href='"&beforepath&"'"
else
word="Finance Series of "&seriesname&" has existed, please input again."
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