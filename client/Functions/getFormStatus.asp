<%
function getFormStatus(thisstatus)
	'0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=deny to transact
	select case thisstatus
	case "0"
	getFormStatus="已申请"
	case "1"
	getFormStatus="已提交"
	case "2"
	getFormStatus="已批准"
	case "3"
	getFormStatus="没有被批准"
	case "4"
	getFormStatus="已处理"
	case "5"
	getFormStatus="所有核准完成"
	case "6"
	getFormStatus="拒绝处理"
	end select
end function
%>
