// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (!shift_type[0].checked&&!shift_type[1].checked)
		{
		alert("Shift Type cannot be blank!����ѡ�������ͣ�");
		return false;
		}
		if (day.value=="")
		{
		alert("Scheduled Day cannot be blank!�ƻ����ڲ���Ϊ�գ�");
		return false;
		}
		if (hour.selectedIndex==0||minute.selectedIndex==0)
		{
		alert("Hour/minute cannot be blank!����ѡ��ƻ���ʱ�䣡")
		return false;
		}
		if (!confirm("Are you sure the time is: " + day.value + " " + hour.options[hour.options.selectedIndex].value + ":" + minute.options[minute.options.selectedIndex].value + "\nȷ�ϼƻ�ʱ���ǣ� " + day.value + " " + hour.options[hour.options.selectedIndex].value + ":" + minute.options[minute.options.selectedIndex].value))
		{
		return false;
		}
	}
}