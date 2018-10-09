<%
const firstpage=1
if request("recordsize")="" then
	recordsize=10
elseif request("recordsize")="all" then
	recordsize=1000
else
	recordsize=cint(request("recordsize"))
end if
if request("pagenum")="" then
	pagenum=1
else
	pagenum=cint(request("pagenum"))
end if
session("pagespan")=request("pagespan")
if session("pagespan")="" then
session("pagespan")=0
end if
%>