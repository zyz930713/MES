// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (code.value=="")
		{
		alert("Code cannot be blank!")
		return false
		}
		if (name.value=="")
		{
		alert("Name cannot be blank!");
		return false;
		}
		if (NT.value=="")
		{
		alert("NT Account cannot be blank!");
		return false;
		}
		if (manager.options[manager.selectedIndex].value==code.value)
		{
		alert("Manager cannot be the same as User code!");
		return false;
		}
		if (factoryto.length==0)
		{
		alert("Factory cannot be blank!");
		return false;
		}
		else
		{
			if (factoryto.length!=0)
			{
			new_toitem_all(document.all.factoryto);
			rolescount.value=factoryto.length;
			}
		}
		if (toitem.length!=0)
		{
		new_toitem_all(document.all.toitem);
		rolescount.value=toitem.length;
		}
	}
}