<%
function getScrapStatus(scrap_status)
	select case scrap_status
	case "0" 'Saved and aprroved
	getScrapStatus="�ѱ���"
	case "1" 'wait to Approve
	getScrapStatus="�ȴ���׼"
	case "2" 'rejected
	getScrapStatus="�����"
	end select
end function
%>