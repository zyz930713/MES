<%
if request("recordsize")="" then
recordsize=10
elseif request("recordsize")="all" then
recordsize=rs.recordcount
else
recordsize=cint(request("recordsize"))
end if
const firstpage=1

rs.pagesize=recordsize
session("strpagenum")=request("pagenum")
if session("strpagenum")="" then
session("strpagenum")=1
end if
session("pagespan")=request("pagespan")
if session("pagespan")="" then
session("pagespan")=0
end if
totalpages=rs.pagecount
%>