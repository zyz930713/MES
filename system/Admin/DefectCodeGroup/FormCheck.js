// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (group_name.value=="")
		{
		alert("Group Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		toitem_all();
	}
}