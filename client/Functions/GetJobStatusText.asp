<%
function CmmsJobStatusText(job_status)
retevl=""
select case job_status
	case "0" 'opened
		retevl="�ȴ�"
	case "1" 'acted
		retevl="���ڴ���"
	case "2" 'canceled
		retevl="��ȡ��"
	case "3" 'closed
		retevl="ȷ�Ϲر�"
	case "4" 'suspended
		retevl="���ݻ�"
	case "5" 'transferred
		retevl="��ת��"
	case "6" 'TEMP_closed
		retevl="��ʱ�ر�"
end select
CmmsJobStatusText=retevl
end function
%>
