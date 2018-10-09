// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(lead_time_unit.selectedIndex==0)
		{
			alert("Unit cannot be blank!")
		}
	}
}