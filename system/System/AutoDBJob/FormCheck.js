// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (day.value=="")
		{
		alert("Scheduled Day cannot be blank!计划日期不得为空！");
		return false;
		}
		if (hour.selectedIndex==0||minute.selectedIndex==0)
		{
		alert("Hour/minute cannot be blank!必须选择计划的时间！")
		return false;
		}
		if (!confirm("Are you sure the time is: " + day.value + " " + hour.options[hour.options.selectedIndex].value + ":" + minute.options[minute.options.selectedIndex].value + "\n确认计划时间是： " + day.value + " " + hour.options[hour.options.selectedIndex].value + ":" + minute.options[minute.options.selectedIndex].value))
		{
		return false;
		}
	}
}