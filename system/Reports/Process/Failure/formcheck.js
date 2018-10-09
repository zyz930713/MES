// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(profile_task_id.selectedIndex==0||seriesgroup.selectedIndex==0)
		{
			alert("Select task and family!");
			return false;
		}
	}
}