<%
function trackJobMaster(job_number,action_type,transact_type,store,inspect,scrap,source_page,row_number)
	html="Job Number:"&job_number&"<br>"
	html=html&"��������:"&action_type&"<br>"
	html=html&"��������:"&transact_type&"<br>"
	html=html&"�ֶ������:"&store&"<br>"
	html=html&"�Զ������:"&inspect&"<br>"
	html=html&"������:"&scrap&"<br>"
	html=html&"����ҳ��:"&source_page&"<br>"
	html=html&"��������:"&row_number&"<br>"
	sendEmail"Form@barcode.com","Dickens.Xu@knowles.com",job_number,html
end function
%>