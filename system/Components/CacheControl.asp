<%
response.Expires=0
response.CacheControl="no-cache"
if trim(session("language")&"")="" then
'response.redirect("/Default.asp")
session("language")=0
end if
if request("path")<>"" then
path=trim(request("path"))
	if request("query")<>"" then
	query=trim(request("query"))
	end if
	if is_return_page=true then
	query=replace(query,"*","&")
	end if
else
path=trim(request.ServerVariables("PATH_INFO"))
query=trim(request.ServerVariables("QUERY_STRING"))
query=replace(query,"&","*")
end if
pagename=trim(request.ServerVariables("PATH_INFO"))
this_path=pagename
beforepath=path&"?"&query
const width="width=""20%"""
const tr_focus="onmouseover=""TR_Over(this)"" onmouseout=""TR_Out(this)"""
if session("langauge")=0 then
username="user"
else
username="cuser"
end if
%>