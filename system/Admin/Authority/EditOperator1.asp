<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
operatorname=trim(request.Form("name"))
SQL="select * from OPERATORS where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
rs("CODE")=trim(request.Form("code"))
rs("OPERATOR_NAME")=operatorname
rs("OPERATOR_CHINESE_NAME")=request.Form("chinesename")
rs("FACTORY_ID")=request.Form("factory")
	if request.Form("locked")="1" then
	rs("LOCKED")=1
	else
	rs("LOCKED")=0
	end if
	if request.Form("practised")="1" then
	rs("PRACTISED")=1
	rs("PRACTISE_START_TIME")=request.Form("fromdate")
	rs("PRACTISE_END_TIME")=request.Form("todate")
	else
	rs("PRACTISED")=0
	rs("PRACTISE_START_TIME")=null
	rs("PRACTISE_END_TIME")=null
	end if
rs.update
word="Successfully edit Operator."
action="location.href='"&beforepath&"'"
else
word="Operator of "&username&" has existed, please input again."
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