<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
groupname=trim(request.Form("groupname"))
roles=replace(request.Form("toitem")," ","")
members=replace(request.Form("toitem2")," ","")
amembers=split(members,",")

errorlist="Following Engineer have added into other groups, please exclude them."
ingroup=false

'for i=0 to ubound(amembers)
'	SQL="select * from USER_GROUP where GROUP_MEMBERS like '%"&amembers(i)&"%'"
'	rs.open SQL,conn,1,3
'	if not rs.eof then
'	ingroup=true
'	errorlist=errorlist&amembers(i)&" in "&rs("GROUP_NAME")&"\n"
'	end if
'	rs.close
'next

updateok=false
if ingroup=false then
	SQL="select * from USER_GROUP where GROUP_NAME='"&groupname&"'"
	rs.open SQL,conn,1,3
	if rs.eof then
	rs.addnew
	rs("NID")="UG"&NID_SEQ("USER_GROUP")
	rs("GROUP_NAME")=groupname
	rs("GROUP_CHINESE_NAME")=trim(request.Form("groupchinesename"))
	rs("FACTORY_ID")=request.Form("factory")
	rs("GROUP_MEMBERS")=members
	rs("ROLES_ID")=roles
	rs("LANGUAGE")=request.Form("language")
	rs.update
	updateok=true
	word="Successfully save a New User Group."
	action="location.href='"&beforepath&"'"
	else
	updateok=false
	word="User Group of "&groupname&" has existed, please input again."
	action="history.back()"
	end if
	rs.close
	if updateok=true and request.Form("apply")="1" then
		for i=0 to ubound(amembers)
		old_roles=""
			new_roles=""
			SQL="select ROLES_ID from USER_GROUP where NID<>'"&id&"' and GROUP_MEMBERS like '%"&amembers(i)&"%'"
			rs.open SQL,conn,1,3
			if not rs.eof then
			while not rs.eof
				if rs("ROLES_ID")<>"" then
				a_role_ids=split(rs("ROLES_ID"),",")
					for j=0 to ubound(a_role_ids)
						if instr(old_roles,a_role_ids(j))<=0 then
						old_roles=old_roles&a_role_ids(j)&","
						end if
					next
				end if
			rs.movenext
			wend
			end if
			rs.close
			if old_roles<>"" then
				if roles<>"" then
				new_roles=old_roles
				a_roles=split(roles,",")
					for j=0 to ubound(a_roles)
						if instr(old_roles,a_roles(j))<=0 then
						new_roles=new_roles&a_roles(j)&","
						end if
					next
				else
				new_roles=left(old_roles,len(old_roles)-1)
				end if
			else
			new_roles=roles
			end if
			SQL="select * from USERS where USER_CODE='"&amembers(i)&"'"
			rs.open SQL,conn,1,3
			if not rs.eof then
			rs("ROLES_ID")=new_roles
			rs("LANGUAGE")=request.Form("language")
			rs.update
			end if
			rs.close
		next
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