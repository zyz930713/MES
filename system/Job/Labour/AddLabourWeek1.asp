<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
line_name=request.Form("line")
year_index=request.Form("year")
week_index=request.Form("week")
week_fromtime=request.Form("fromdate")&" "&request.Form("fromhour")&":"&request.Form("fromminute")
week_totime=request.Form("todate")&" "&request.Form("tohour")&":"&request.Form("tominute")
SQL="select * from LINE_LABOUR where LINE_NAME='"&line_name&"' and YEAR_INDEX='"&year_index&"' and WEEK_INDEX='"&week_index&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	NID="LA"&NID_SEQ("LABOUR")
	rs("NID")=NID
	rs("LINE_NAME")=line_name
	rs("YEAR_INDEX")=year_index
	rs("WEEK_INDEX")=week_index
	rs("WEEK_FROM_TIME")=week_fromtime
	rs("WEEK_END_TIME")=week_totime
	rs.update
	word=line_name&"线 "&year_index&" 第"&week_index&"周的劳动力新增成功！"
	action="location.href='"&beforepath&"'"
else
word=line_name&"线 "&year_index&" 第"&week_index&"周的劳动力已存在！"
action="history.back()"
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->