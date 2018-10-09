<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="Retest_Check.asp"
nid=trim(request.QueryString("nid"))
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
'retest check in job_master_store
if retestchecker=true then
SQL="update JOB_MASTER_STORE set RETEST_CHECK_STATUS=1,RETEST_CHECK_CODE='"&session("code")&"',RETEST_CHECK_TIME=to_date('"&now()&"','yyyy-mm-dd hh24:mi:ss') where NID='"&nid&"'"
rs.open SQL,conn,1,3
word="复查成功！"
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