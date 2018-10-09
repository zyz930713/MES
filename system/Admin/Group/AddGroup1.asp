<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Group/GroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
groupname=trim(request.Form("groupname"))
members=replace(request.Form("members")," ","")
amembers=split(members,",")

errorlist="Following Jobs have added into other groups, please exclude them."
ingroup=false

for i=0 to ubound(amembers)
	SQL="select * from SYSTEM_GROUP where GROUP_MEMBERS like '%"&amembers(i)&"%'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	ingroup=true
	errorlist=errorlist&amembers(i)&" in "&rs("GROUP_NAME")&"\n"
	end if
	rs.close
next

if ingroup=false then
	SQL="select * from SYSTEM_GROUP where GROUP_NAME='"&groupname&"'"
	rs.open SQL,conn,1,3
	if rs.eof then
	rs.addnew
	rs("NID")="GR"&NID_SEQ("GROUP")
	rs("GROUP_NAME")=groupname
	rs("GROUP_CHINESE_NAME")=trim(request.Form("groupchinesename"))
	rs("FACTORY_ID")=request.Form("factory")
	rs("GROUP_TYPE")=request.Form("grouptype")
	rs("GROUP_MEMBERS")=members
	rs.update
	rs.close
	word="Successfully save a New Group."
	action="location.href='"&beforepath&"'"
	else
	word="Group of "&groupname&" has existed, please input again."
	action="history.back()"
	end if
else
	word=errorlist
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