// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(line.selectedIndex==0)
		{
		alert("线别不得为空！");
		return false;
		}
		if(fromdate.value=="")
		{
		alert("必须选择周次开始时间！");
		return false;
		}
		if(todate.value=="")
		{
		alert("必须选择周次结束时间！");
		return false;
		}
		if(fromhour.value=="")
		{
		alert("周次开始时间不完整！");
		return false;
		}
		if(fromminute.value=="")
		{
		alert("周次开始时间不完整！");
		return false;
		}
		if(tohour.value=="")
		{
		alert("周次结束时间不完整！");
		return false;
		}
		if(tominute.value=="")
		{
		alert("周次结束时间不完整！");
		return false;
		}
		if (!(isNumberString(year.value,"1234567890"))||year.value.length!=4)
		{
		alert("年度格式错误！");
		return false;
		}	
		if (!(isNumberString(week.value,"1234567890"))||new Number(week.value)>52)
		{
		alert("周次格式错误！");
		return false;
		}	
	}
}