<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
code=trim(request.Form("code"))
operatorname=trim(request.Form("name"))
SQL="select * from OPERATORS where CODE='"&code&"'"
rs.open SQL,conn,3,3
if rs.eof then
rs.addnew
rs("NID")="OP"&NID_SEQ("OPERATOR")
rs("CODE")=code
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
	rs("PRACTISE_START_TIME")=request.Form("fromtime")
	rs("PRACTISE_END_TIME")=request.Form("totime")
	else
	rs("PRACTISED")=0
	rs("PRACTISE_START_TIME")=null
	rs("PRACTISE_END_TIME")=null
	end if
rs.update
word="Successfully save New Operator."
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