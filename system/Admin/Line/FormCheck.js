// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (linename.value=="")
		{
		alert("Line Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (section.selectedIndex==0)
		{
		alert("Section cannot be blank!");
		return false;
		}
	}
}