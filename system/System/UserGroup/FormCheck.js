// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (groupname.value=="")
		{
		alert("Group Name cannot be blank!");
		return false;
		}
		if (groupchinesename.value=="")
		{
		alert("Group Chinese Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (language.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (toitem.length!=0)
		{
		toitem_all();
		}
		if (toitem2.length!=0)
		{
		toitem2_all();
		}
	}
}