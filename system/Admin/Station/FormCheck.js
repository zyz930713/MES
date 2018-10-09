// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (stationnumber.value=="")
		{
		alert("Station Number cannot be blank!");
		return false;
		}
		if (stationname.value=="")
		{
		alert("Station Name cannot be blank!");
		return false;
		}
		if (stationchinesename.value=="")
		{
		alert("Station Chinese Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (section.selectedIndex==0)
		{
		alert("Section cannot be blank!")
		return false;
		}
		if (quantity_type[0].checked==false&&quantity_type[1].checked==false)
		{
		alert("Intial Quantity Type cannot be blank!");
		return false;
		}
		
		
		if (toitem2.length!=0)
		{
		toitem2_all();
		//actionscount.value=toitem2.length;
		}
		
	}
}