<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="DeleteStore.asp"
nid=trim(request.QueryString("id"))
factory=trim(request.QueryString("factory"))
valid=false
diff=0
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
'delete store info in job_master_store
SQL="select * from JOB_MASTER_STORE where NID='"&nid&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("FACTORY_ID")=factory
rs.update
valid=true
word="更新成功！"
end if
rs.close
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