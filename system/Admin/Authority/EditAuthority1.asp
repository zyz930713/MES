<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.form("path")
query=request.form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
icount=request.Form("icount")
jcount=request.Form("jcount")
code=request.Form("code")
nids=""
pids=""
for i=1 to icount
	if request.Form("nid"&i)="1" then
	nids=nids&request.Form("n_nid"&i)&","
	end if
next
for j=1 to jcount
	if request.Form("pid"&j)="1" then
	pids=pids&request.Form("p_pid"&j)&","
	end if
next
if nids<>"" then
nids=left(nids,len(nids)-1)
end if
if pids<>"" then
pids=left(pids,len(pids)-1)
end if
SQL="Select AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID from OPERATORS where code='"&code&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("AUTHORIZED_STATIONS_ID")=nids
rs("AUTHORIZED_PARTS_ID")=pids
rs.update
end if
rs.close
word="Suscessfully update "&code&"'s authority"
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>
</script>
<body>

</body>
</html>
