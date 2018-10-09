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
		if (defectname.selectedIndex==0)
		{
		alert("Defect Name cannot be blank!")
		return false;
		}
		if (toitem.length==0)
		{
		alert("Applied Materials cannot be blank!");
		return false;
		}
		else
		{
		toitem_all();
		materialscount.value=toitem.length;
		}
		if (station.selectedIndex==0)
		{
		alert("Belonged Station cannot be blank!")
		return false;
		}
	}
}