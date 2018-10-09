<%

pagesize_s=request.Form("select")
pagesize_a=request.querystring("pagesize_a")
select case pagesize_s
case "10" 
pagesize_s=10
case "15" 
pagesize_s=15
case "20" 
pagesize_s=20
case "25" 
pagesize_s=25
case "30" 
pagesize_s=30
case "all" 
pagesize_s=rsPR.recordcount
case else 
	if pagesize_a="" then
	pagesize_s=10
	else 
	pagesize_s=pagesize_a
	end if
end select
const firstpage=1

rsPR.pagesize=pagesize_s
totalpages=rsPR.pagecount
session("strpagenum")=request.querystring("pagenum")
if session("strpagenum")="" then
session("strpagenum")=1
end if
session("pagespan")=request.querystring("pagespan")
if session("pagespan")="" then
session("pagespan")=0
end if
%>
