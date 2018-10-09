// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (seriesgroupname.value=="")
		{
			alert("Family Name cannot be blank!");
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
		if (first_yield.value=="")
		{
			alert("First Yield cannot be blank!");
			return false;
		}
		if (internal_yield.value=="")
		{
			alert("Internal Yield cannot be blank!");
			return false;
		}
		if (final_yield.value=="")
		{
			alert("Target Yield cannot be blank!");
			return false;
		}
		if (inspect_yield.value=="")
		{
			alert("Target of Retest Yield  cannot be blank!");
			return false;
		}
		toitem_all();
	}
}