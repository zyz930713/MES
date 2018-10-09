// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (machinenumber.value=="")
		{
		alert("Machine Number cannot be blank!");
		return false;
		}
		if (machinename.value=="")
		{
		alert("Machine Name cannot be blank!")
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (toitem.length!=0)
		{
		toitem_all();
		stationscount.value=toitem.length;
		}

	}
}