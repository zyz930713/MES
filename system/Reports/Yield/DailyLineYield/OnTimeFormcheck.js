// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(factory_id.selectedIndex==0)
		{
			alert("Factory Name cannot be blank!");
			return false;
		}
	}
}