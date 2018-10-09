<%
if admin<>true then
response.Redirect("/main.asp?error=Error message: you are not authorized to access this page.")
end if
%>