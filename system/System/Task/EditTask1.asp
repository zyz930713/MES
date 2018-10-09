<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
taskname=trim(request.Form("taskname"))
SQL="select * from TASK where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
rs("TASK_NAME")=taskname
rs("TASK_CHINESE_NAME")=trim(request.Form("taskchinesename"))
rs("DESCRIPTION")=trim(request.Form("description"))
rs("PACKAGE")=trim(request.Form("package"))
if request.Form("apply_by_user")="1" then
rs("APPLY_BY_USER")=1
else
rs("APPLY_BY_USER")=0
end if
rs("PARAM1")=trim(request.Form("param1"))
rs("PARAM_CHINESE1")=trim(request.Form("cparam1"))
rs("PARAM_TYPE1")=trim(request.Form("paramtype1"))
rs("PARAM2")=trim(request.Form("param2"))
rs("PARAM_CHINESE2")=trim(request.Form("cparam2"))
rs("PARAM_TYPE2")=trim(request.Form("paramtype2"))
rs("PARAM3")=trim(request.Form("param3"))
rs("PARAM_CHINESE3")=trim(request.Form("cparam3"))
rs("PARAM_TYPE3")=trim(request.Form("paramtype3"))
rs("PARAM4")=trim(request.Form("param4"))
rs("PARAM_CHINESE4")=trim(request.Form("cparam4"))
rs("PARAM_TYPE4")=trim(request.Form("paramtype4"))
rs.update
word="Successfully edit Task."
action="location.href='"&beforepath&"'"
else
word="Task of "&taskname&" has existed, please input again."
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