// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (ITEM_NAME.value=="")
		{
		alert("12NC Name cannot be blank!");
		ITEM_NAME.focus();
		return false;
		}
		if (SUPPLIER_NAME.value=="")
		{
		alert("SUPPLIER NAME  cannot be blank!");
		SUPPLIER_NAME.focus();
		return false;
		}
		if (STATION_DESCRIPTION.value=="")
		{
		alert("STATION DESCRIPTION cannot be blank!");
		STATION_DESCRIPTION.focus();
		return false;
		}
		if (STATION_DESCRIPTION_EN.value=="")
		{
		alert("project  cannot be blank!");
		STATION_DESCRIPTION_EN.focus();
		return false;
		}
		
	}
}