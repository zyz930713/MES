function formcheck()
{
	with(document.form1)
	{
		if(englishname=="")
		{
		alert("English Name cannot be blank!英文名不能为空！");
		return false;
		}
		if(chinesename=="")
		{
		alert("Chinese Name cannot be blank!中文名不能为空！");
		return false;
		}
		if(language.selectedIndex==0)
		{
		alert("Langauge cannot be blank!语言不能为空！");
		return false;
		}
		for(var i=1;i<=15;i++)
		{
			if(eval("prefix"+i+".value!=''"))
			{
				if(eval("model_yield"+i+".value==''"))	
				{
				alert("Model Yield for "+eval("prefix"+i+".value")+" cannot be blank!");
				return false;
				}
			}
			if(eval("line"+i+".value!=''"))
			{
				if(eval("line_yield"+i+".value==''"))	
				{
				alert("Line Yield for "+eval("line"+i+".value")+" cannot be blank!");
				return false;
				}
			}
			if(eval("family"+i+".selectedIndex!=0"))
			{
				if(eval("family_yield"+i+".value==''"))	
				{
				alert("Family Yield for "+eval("family"+i+".options[family"+i+".selectedIndex].text")+" cannot be blank!");
				return false;
				}
			}
		}
	}
}