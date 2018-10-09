// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (materialnumber.value=="")
		{
		alert("Material Number cannot be blank!");
		return false;
		}
		if (materialname.value=="")
		{
		alert("Material Name cannot be blank!")
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (materialunit.selectedIndex==0)
		{
		alert("Material Unit cannot be blank!")
		return false;
		}
	}
}