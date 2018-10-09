<%
function CompareLeadtime(part_nummber,day_time,job_time) 
set rss=server.CreateObject("adodb.recordset")
'SQL="select LEAD_TIME from SERIES_GROUP where instr(INCLUDED_SYSTEM_ITEMS,'"&part_nummber&"')>0"
SQL="select lead_time,lead_time2 from product_model where item_name='"&part_nummber&"' "
rss.open SQL,conn,1,3
if not rss.eof then
	leadtime=csng(rss("lead_time"))
	if job_time>leadtime then
		CompareLeadtime="<span class=""red"">"&day_time&" ("&leadtime/60/24&")</span>"
	else
		CompareLeadtime=day_time&" ("&leadtime/60/24&")"
	end if
	if rss("lead_time2") <> "0" then
		CompareLeadtime=CompareLeadtime&"("&rss("lead_time2")/60/24&")"
	end if
end if
rss.close
end function

function CompareLeadtimeExcel(part_nummber,day_time,job_time,red) 
red=false
set rss=server.CreateObject("adodb.recordset")
'SQL="select LEAD_TIME from SERIES_GROUP where instr(INCLUDED_SYSTEM_ITEMS,'"&part_nummber&"')>0"
SQL="select lead_time,lead_time2 from product_model where item_name='"&part_nummber&"' "
rss.open SQL,conn,1,3
if not rss.eof then
	leadtime=csng(rss("lead_time"))
	if job_time>leadtime then
		red=true
	end if
	CompareLeadtimeExcel=day_time&" ("&leadtime/60/24&")"
	if rss("lead_time2") <> "0" then
		CompareLeadtimeExcel=CompareLeadtimeExcel&"("&rss("lead_time2")/60/24&")"
	end if
end if
rss.close
end function
%>