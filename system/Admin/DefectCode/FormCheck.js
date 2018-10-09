// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (defectcode.value=="")
		{
		alert("Defect Code cannot be blank!");
		return false;
		}
		if (defectname.value=="")
		{
		alert("Defect Name cannot be blank!");
		return false;
		}
		if (chinesename.value=="")
		{
		alert("Defect Chinese Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!");
		return false;
		}
		if (toitem1.length==0)
		{
		alert("Applied Materials cannot be blank!");
		return false;
		}
		else
		{
		toitem1_all();
		materialscount.value=toitem1.length;
		}
		toitem2_all();
		if (station.selectedIndex==0)
		{
		alert("Belonged Station cannot be blank!")
		return false;
		}
	}
}
function formcheck_New()
{
	with(document.form1)
	{
		if (defectcode.value=="")
		{
		alert("Defect Code cannot be blank!");
		return false;
		}
		if (defectname.value=="")
		{
		alert("Defect Name cannot be blank!");
		return false;
		}
		if (chinesename.value=="")
		{
		alert("Defect Chinese Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!");
		return false;
		}		
	}
}