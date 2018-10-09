// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		try
		{
			if (thisform.selectedIndex==0)
			{
			alert("Form cannot be blank!\n必须选择一个表单！")
			return false;
			}
		}
		catch(e){};
		try
		{
			if (param1.value=="")
			{
			alert("Param 1 cannot be blank!\n参数1不得为空！");
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
				alert("Param 2 cannot be blank!\n参数2不得为空！");
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
				alert("Param 2 cannot be blank!\n参数2不得为空！");
				return false;
				}
			}
		}
		catch(e){};
		try
		{
			if (param3.value=="")
			{
			alert("Param 3 cannot be blank!\n参数3不得为空！");
			return false;
			}
		}
		catch(e){};
	}
}