<%
function trackJobMaster(job_number,action_type,transact_type,store,inspect,scrap,source_page,row_number)
	html="Job Number:"&job_number&"<br>"
	html=html&"操作类型:"&action_type&"<br>"
	html=html&"处理类型:"&transact_type&"<br>"
	html=html&"手动入库数:"&store&"<br>"
	html=html&"自动入库数:"&inspect&"<br>"
	html=html&"报废数:"&scrap&"<br>"
	html=html&"处理页面:"&source_page&"<br>"
	html=html&"代码行数:"&row_number&"<br>"
	sendEmail"Form@barcode.com","Dickens.Xu@knowles.com",job_number,html
end function
%>