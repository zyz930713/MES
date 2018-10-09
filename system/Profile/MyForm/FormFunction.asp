<%
function getFormStatus(thisstatus)
	'0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=deny to transact
	if session("language")="0" then
		select case thisstatus
		case "0"
		getFormStatus="Applied"
		case "1"
		getFormStatus="Submited"
		case "2"
		getFormStatus="Approved"
		case "3"
		getFormStatus="Disapproved"
		case "4"
		getFormStatus="Transacted"
		case "5"
		getFormStatus="Approval Finished"
		case "6"
		getFormStatus="Rejected"
		end select
	else
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
	end if
end function

function getFormType(thistype)
	'0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=deny to transact
	if session("language")="0" then
		select case thistype
		case "0"
		getFormType="Not need to be approved"
		case "1"
		getFormType="Need to be approved"
		end select
	else
		select case thistype
		case "0"
		getFormType="不需要核准"
		case "1"
		getFormType="需要核准"
		end select
	end if
end function

function getParamName(param,param_chinese)
	if param<>"" then
		if session("language")="0" then
		getParamName="&nbsp;("&param&")"
		else
		getParamName="&nbsp;("&param_chinese&")"
		end if
	end if
end function
%>
