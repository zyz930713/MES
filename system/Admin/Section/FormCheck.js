// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (sectionname.value=="")
		{
		alert("Section Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
	}
}