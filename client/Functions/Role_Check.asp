<%
Page_Role = session("Page_Role")
if instr(session("role"),Page_Role) = 0 then
	response.write "<script>alert('���ӳ�ʱ,������û�в�����ҳ���Ȩ��!');parent.location.href='../../Functions/user_login.asp'</script>"
	response.end()
end if
%>