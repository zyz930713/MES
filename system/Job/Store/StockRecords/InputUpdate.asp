<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsStockAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="Retest_Check.asp"
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
nid=trim(request.QueryString("nid"))
new_input=cint(request.QueryString("new_input"))
if checker=true then
	SQL="select INPUT_QUANTITY,STORE_QUANTITY,YIELD from JOB_MASTER_STORE where NID='"&nid&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	rs("INPUT_QUANTITY")=new_input
	rs("YIELD")=cint(rs("STORE_QUANTITY"))/new_input
	rs.update
	end if
	rs.close
word="更新成功！"
else
word="无权操作！"
end if
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->