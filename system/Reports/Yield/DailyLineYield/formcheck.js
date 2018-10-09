// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(profile_task_id.selectedIndex==0)
		{
			alert("Task Name cannot be blank!");
			return false;
		}
		if(factory_id.selectedIndex==0)
		{
			alert("Factory Name cannot be blank!");
			return false;
		}
	}
}