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
		if(week_number.value=="")
		{
			alert("Week Number cannot be blank!");
			return false;
		}
		if(year_number.value=="")
		{
			alert("Year Number cannot be blank!");
			return false;
		}
	}
}