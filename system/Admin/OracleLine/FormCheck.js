// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(isNumberString(subqty.value,'1234567890')!=1)
		{
			alert("Number format is error!");
			return false;
		}
	}
}