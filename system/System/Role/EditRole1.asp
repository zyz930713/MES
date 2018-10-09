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
rolename=trim(request.Form("name"))
members=trim(request.Form("members"))
members=replace(members," ","")
old_memebers=trim(request.Form("old_memebers"))
if old_memebers<>"" and members<>"" then
a_old_memebers=split(old_memebers,",")
canceled_members=""
	for i=0 to ubound(a_old_memebers)
		if instr(members,a_old_memebers(i))<=0 then
		canceled_members=canceled_members&a_old_memebers(i)&","
		end if
	next
else
canceled_members=old_memebers&","
end if
SQL="select * from ENGINEER_ROLE where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
old_role_name=trim(rs("ROLE_NAME"))
rs("ROLE_NAME")=rolename
rs("ROLE_CHINESE_NAME")=trim(request.Form("chinesename"))
rs("DESCRIPTION")=trim(request.Form("description"))
rs("APPLICANT")=trim(request.Form("applicant"))
if request.Form("apply")="1" then
rs("APPLY_FROM_WEB")=1
else
rs("APPLY_FROM_WEB")=0
end if
rs.update
rs.close
update_count=0
cancel_count=0
if members<>"" and old_role_name<>rolename then
a_members=split(members,",")
	for i=0 to ubound(a_members)
	SQL="select * from USERS where USER_CODE='"&a_members(i)&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	rs("ROLES_ID")=replace(rs("ROLES_ID"),old_role_name,rolename)
	update_count=update_count+1
	rs.update
	end if
	rs.close
	next
end if
if canceled_members<>"" then
canceled_members=left(canceled_members,len(canceled_members)-1)
a_canceled_members=split(canceled_members,",")
	for i=0 to ubound(a_canceled_members)
	SQL="select * from USERS where USER_CODE='"&a_canceled_members(i)&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		if rs("ROLES_ID")<>"" then
			a_roles_id=split(rs("ROLES_ID"),",")
			new_role_id=""
			for j=0 to ubound(a_roles_id)
				if a_roles_id(j)<>old_role_name then
				new_role_id=new_role_id&","
				end if
			next
			if new_role_id<>"" then
			new_role_id=left(new_role_id,len(new_role_id)-1)
			rs("ROLES_ID")=new_role_id
			else
			rs("ROLES_ID")=null
			end if
			cancel_count=cancel_count+1
			rs.update
		end if
	end if
	rs.close
	next
end if
word="Successfully edit Role. Update "&update_count&"; Cancel "&cancel_count
action="location.href='"&beforepath&"'"
else
word="Role of "&rolename&" has existed, please input again."
action="history.back()"
rs.close
end if
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