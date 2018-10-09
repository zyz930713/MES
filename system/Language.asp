<%
session("language")=cint(request("language"))
response.Redirect("/Default.asp")
%>
