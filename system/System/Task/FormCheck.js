// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (taskname.value=="")
		{
		alert("Task Name cannot be blank!");
		return false;
		}
	}
}