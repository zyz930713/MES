// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!");
		return false;
		}
	}
}