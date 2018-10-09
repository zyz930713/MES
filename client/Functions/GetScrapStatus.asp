<%
function getScrapStatus(scrap_status)
	select case scrap_status
	case "0" 'Saved and aprroved
	getScrapStatus="已保存"
	case "1" 'wait to Approve
	getScrapStatus="等待核准"
	case "2" 'rejected
	getScrapStatus="被否决"
	end select
end function
%>