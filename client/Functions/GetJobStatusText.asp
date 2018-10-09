<%
function CmmsJobStatusText(job_status)
retevl=""
select case job_status
	case "0" 'opened
		retevl="等待"
	case "1" 'acted
		retevl="正在处理"
	case "2" 'canceled
		retevl="被取消"
	case "3" 'closed
		retevl="确认关闭"
	case "4" 'suspended
		retevl="被暂缓"
	case "5" 'transferred
		retevl="被转移"
	case "6" 'TEMP_closed
		retevl="暂时关闭"
end select
CmmsJobStatusText=retevl
end function
%>
