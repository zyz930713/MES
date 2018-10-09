// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (action[0].checked==false&&action[1].checked==false)
		{
		alert("Approve Type cannot be blank!\n核准方式不得为空！");
		return false;
		}
		if (action[1].checked==true&&denyreason.value=="")
		{
		alert("Deny Reason cannot be blank!\n否决理由不得为空！");
		return false;
		}
	}
}