// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (name.value=="")
		{
		alert("Name cannot be blank!");
		return false;
		}
		if (chinesename.value=="")
		{
		alert("Chinese Name cannot be blank!");
		return false;
		}
	}
}