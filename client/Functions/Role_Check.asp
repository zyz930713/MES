<%
Page_Role = session("Page_Role")
if instr(session("role"),Page_Role) = 0 then
	response.write "<script>alert('连接超时,或者您没有操作此页面的权限!');parent.location.href='../../Functions/user_login.asp'</script>"
	response.end()
end if
%>