// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (toitem.length==0)
		{
		alert("Selected Stations cannot be blank!");
		return false;
		}
		else
		{
		fromitem_all();
		toitem_all();
		operatorcount.value=toitem.length;
		}
	}
}