// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (action[0].checked==false&&action[1].checked==false)
		{
		alert("Transact Type cannot be blank!\n处理方式不得为空！");
		return false;
		}
		if (action[1].checked==true&&rejectreason.value=="")
		{
		alert("Reject Reason cannot be blank!\n拒绝理由不得为空！");
		return false;
		}
	}
}