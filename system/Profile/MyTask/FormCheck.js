// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		try{
			if (task.selectedIndex==0)
			{
			alert("Task cannot be blank!\n����ѡ��һ������")
			return false;
			}
		}
		catch(e){}
		if (taskname.value=="")
		{
		alert("Task Name cannot be blank!\n�������Ʋ���Ϊ�գ�");
		return false;
		}
		if (toitem.length!=0)
		{
		toitem_all();
		}
		else
		{
		alert("Must select at least one reciever, such as yourself!\n����ѡ��һ���ռ��ˣ����Լ���");
		return false;
		}
		if(period.checked)
		{
			if(fromdate.value=='')
			{
				alert("Must set starting time.\n�������üƻ���ʼʱ�䣡");
				return false;
			}
			if(happenitem.selectedIndex==0)
			{
				alert("Must set schedule period.\n�������üƻ����ڣ�");
				return false;
			}
			if(happenitem.selectedIndex==3)
			{
				if(week[0].checked==false&&week[1].checked==false&&week[2].checked==false&&week[3].checked==false&&week[4].checked==false&&week[5].checked==false&&week[6].checked==false)
				{	
				alert("Must set schedule week day.\n�������üƻ����������ڣ�");
				return false;
				}
			}
		}
		if(period.checked==false)
		{
			if(confirm("You has not set schedule. Would you tun task at once?"))
			{
			atonce.value="1";
			}
			else
			{
			atonce.value="0";
			}
		}
		else
		{
			atonce.value="0";
		}
	}
}

function getParam()
{
	with(document.form1)
	{
		if(task.selectedIndex!=0)
		{
			document.all.paramFrame.src="/Functions/getTaskParam.asp?id="+task.options[task.selectedIndex].value;
		}
		else
		{
			document.all.paramhtml1.innerHTML="";
			document.all.paramhtml2.innerHTML="";
		}
	}
}
function isperiod()
{
	with(document.form1)
	{
		if(period.checked)
		{
			happenitem.disabled=false;
			fromdate.disabled=false;
			happentime1.disabled=false;
			happentime2.disabled=false;
		}
		else
		{
			happenitem.disabled=true;
			fromdate.disabled=true;
			happentime1.disabled=true;
			happentime2.disabled=true;
			happenitem.options[0].selected=true;
		}
		isweek();
	}
}
function isweek()
{
	with(document.form1)
	{
		if(happenitem.selectedIndex==3)
		{
			span_hour.style.visibility="hidden";
			span_week.style.visibility="visible";
			week[0].disabled=false;
			week[1].disabled=false;
			week[2].disabled=false;
			week[3].disabled=false;
			week[4].disabled=false;
			week[5].disabled=false;
			week[6].disabled=false;
		}
		else if(happenitem.selectedIndex==1)
		{
			span_hour.style.visibility="visible";
			span_week.style.visibility="hidden";
			week[0].disabled=true;
			week[1].disabled=true;
			week[2].disabled=true;
			week[3].disabled=true;
			week[4].disabled=true;
			week[5].disabled=true;
			week[6].disabled=true;
		}
		else
		{
			span_hour.style.visibility="hidden";
			span_week.style.visibility="hidden";
		}
	}
}

function hourcheck(ob)
{
	with(document.form1)
	{
		if (!(isNumberString(ob.value,"1234567890"))||new Number(ob.value)>24)
		{
			alert("ʱ���ʽ����");
			ob.value="00";
		}	
		else
		{
			if (ob.value=="24")
			{
			ob.value="00";
			}
		}
	}
}

function minutecheck(ob)
{
	with(document.form1)
	{
		if (!(isNumberString(ob.value,"1234567890"))||new Number(ob.value)>60)
		{
			alert("ʱ���ʽ����");
			ob.value="00";
		}	
		else
		{
			if (ob.value=="60")
			{
			ob.value="00";
			}
		}
	}
}