// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (formname.value=="")
		{
		alert("Form Name cannot be blank!");
		return false;
		}
		if (toitem.length!=0)
		{
		toitem_all();
		}
	}
}