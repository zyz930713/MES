<%
function getFormStatus(thisstatus)
	'0=open;1=submit;2=finished all approval;3=fail to be approved;4=transacted;5=deny to transact
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
end function
%>
