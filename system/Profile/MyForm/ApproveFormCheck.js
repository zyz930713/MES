// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (action[0].checked==false&&action[1].checked==false)
		{
		alert("Approve Type cannot be blank!\n��׼��ʽ����Ϊ�գ�");
		return false;
		}
		if (action[1].checked==true&&denyreason.value=="")
		{
		alert("Deny Reason cannot be blank!\n������ɲ���Ϊ�գ�");
		return false;
		}
	}
}