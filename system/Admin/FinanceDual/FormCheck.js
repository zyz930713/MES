// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (Dual.value=="")
		{
		alert("Dual Part Number cannot be blank!");
		return false;
		}
		if (Single1.value=="")
		{
		alert("Single 1 Part Number cannot be blank!");
		return false;
		}
		if (Single2.value=="")
		{
		alert("Single 2 Part Number cannot be blank!");
		return false;
		}
	}
}