// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (factoryname.value=="")
		{
		alert("Factory Name cannot be blank!");
		return false;
		}
	}
}