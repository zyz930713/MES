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
		getFormStatus="������"
		case "1"
		getFormStatus="���ύ"
		case "2"
		getFormStatus="����׼"
		case "3"
		getFormStatus="û�б���׼"
		case "4"
		getFormStatus="�Ѵ���"
		case "5"
		getFormStatus="���к�׼���"
		case "6"
		getFormStatus="�ܾ�����"
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
		getFormType="����Ҫ��׼"
		case "1"
		getFormType="��Ҫ��׼"
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
