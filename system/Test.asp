<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from OPERATORS order by code"
rs.open SQL,conn,3,3
if not rs.eof then
while not rs.eof
	SQL1="select * from OPERATORS1 where code='"&rs("code")&"'"
	rs1.open SQL1,conn,3,3
	if rs1.eof then
	rs1.addnew
	rs1("NID")="OP"&NID_SEQ("OPERATOR")
	rs1("CODE")=rs("CODE")
	rs1("OPERATOR_NAME")=rs("OPERATOR_NAME")
	rs1("OPERATOR_CHINESE_NAME")=rs("OPERATOR_CHINESE_NAME")
	rs1("TEMP")=rs("TEMP")
	rs1.update
	end if
	rs1.close
rs.movenext
wend
end if
rs.close
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->