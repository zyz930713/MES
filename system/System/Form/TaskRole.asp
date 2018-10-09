<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.QueryString("id")
rolename=request.QueryString("rolename")
SQL="delete from ENGINEER_ROLE where NID='"&id&"'"
rs.open SQL,conn,3,3
SQL="select ROLES_ID from USERS where ROLES_ID like '%"&rolename&"%'"
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
		new_roles_id=replace(rs("ROLES_ID"),rolename&",","")
		if right(new_roles_id,1)="," then
		new_roles_id=left(new_roles_id,len(new_roles_id)-1)
		end if
		rs("ROLES_ID")=new_roles_id
	rs.update
	rs.movenext
	wend
end if
rs.close
word="Role of "&rolename&" is deleted forever!"
action="location.href='"&beforepath&"'"
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
