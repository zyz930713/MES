// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (seriesgroupname.value=="")
		{
		alert("Series Group Name cannot be blank!");
		return false;
		}
		if (prefix.value=="")
		{
		alert("Prefix cannot be blank!");
		return false;
		}
		if (yield.value=="")
		{
		alert("Yield cannot be blank!");
		return false;
		}
		toitem_all();
	}
}