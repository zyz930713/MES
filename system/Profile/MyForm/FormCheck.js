// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		try
		{
			if (thisform.selectedIndex==0)
			{
			alert("Form cannot be blank!\n����ѡ��һ������")
			return false;
			}
		}
		catch(e){};
		try
		{
			if (param1.value=="")
			{
			alert("Param 1 cannot be blank!\n����1����Ϊ�գ�");
			return false;
			}
		}
		catch(e){};
		try
		{
			if (paramname2.value=="New Quantity")
			{
				if (param2.value=="")
				{
				alert("Param 2 cannot be blank!\n����2����Ϊ�գ�");
				return false;
				}
				if (compareJobQuantity()==false)
				{
					return false;
				}
			}
			else if(paramname2.value=="New Line")
			{
				if (param2.selectedIndex==0)
				{
				alert("Param 2 cannot be blank!\n����2����Ϊ�գ�");
				return false;
				}
			}
		}
		catch(e){};
		try
		{
			if (param3.value=="")
			{
			alert("Param 3 cannot be blank!\n����3����Ϊ�գ�");
			return false;
			}
		}
		catch(e){};
	}
}