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
	}
}