// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if(line.selectedIndex==0)
		{
		alert("�߱𲻵�Ϊ�գ�");
		return false;
		}
		if(fromdate.value=="")
		{
		alert("����ѡ���ܴο�ʼʱ�䣡");
		return false;
		}
		if(todate.value=="")
		{
		alert("����ѡ���ܴν���ʱ�䣡");
		return false;
		}
		if(fromhour.value=="")
		{
		alert("�ܴο�ʼʱ�䲻������");
		return false;
		}
		if(fromminute.value=="")
		{
		alert("�ܴο�ʼʱ�䲻������");
		return false;
		}
		if(tohour.value=="")
		{
		alert("�ܴν���ʱ�䲻������");
		return false;
		}
		if(tominute.value=="")
		{
		alert("�ܴν���ʱ�䲻������");
		return false;
		}
		if (!(isNumberString(year.value,"1234567890"))||year.value.length!=4)
		{
		alert("��ȸ�ʽ����");
		return false;
		}	
		if (!(isNumberString(week.value,"1234567890"))||new Number(week.value)>52)
		{
		alert("�ܴθ�ʽ����");
		return false;
		}	
	}
}