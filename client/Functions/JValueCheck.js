// JavaScript Document
function valuecheck(index,subindex,action_purpose,action_name)
{
	var valid=true;
	var word="";
	var quantity_total=0;
	with (document.form1)
	{
			if (action_purpose=="4")//if it is Material Quantity
			{
				var actNumber = eval("action_number"+index+".value")
				for (var i=1;i<=actNumber;i++)//get total quantity of this kind of material part
				{
					eval("if(action_value"+index+"_"+i+".value!='')")
					{
						quantity_total=quantity_total+eval("new Number(document.form1.action_value"+index+"_"+i+".value)")
					}	
				}
				var max_quantity=eval("new Number(max_quantity"+index+".value)")
				var min_quantity=eval("new Number(min_quantity"+index+".value)")
				if(max_quantity!=0&&max_quantity!="")
				{
					if(quantity_total>max_quantity)
					{
						valid=false;
						word="Total quantity of "+action_name+" exceeds max number!\n"+action_name+"的数量总和超过最大值！";
						object=eval("action_value"+index+"_"+subindex);
					}
				}
				if(min_quantity!=0&&min_quantity!="")
				{
					if(quantity_total<min_quantity)
					{
						valid=false;
						word="Total quantity of "+action_name+" is less than min number!\n"+action_name+"的数量总和小于最小值！";
						object=eval("action_value"+index+"_"+subindex);
					}
				}
			}
			else
			{	
				var validvalue=eval("valid_value"+index+".value")
				var putvalue=eval("action_value"+index+"_"+subindex+".value")
				
				if(validvalue!="")
				{
					valid=true;
					var a_validvalue=new Array();
					a_validvalue=validvalue.split(",");
					forLoop1:
					for (var l=0;l<a_validvalue.length;l++)
					{	
						for(var m=0;m<a_validvalue[l].length;m++)
						{
							var this_validchar=a_validvalue[l].substring(m,m+1);
							var this_putchar=putvalue.substring(m,m+1);
							if(this_validchar!="*"&&this_validchar!="%")
							{
								if(this_validchar!=this_putchar)
								{
									valid=false;
									word="Value of "+action_name+" is invalid!\n"+action_name+"的输入值错误！";
									object=eval("action_value"+index+"_"+subindex);									
									break;	
								}
								else
								{
									valid=true;
								}
							}
							else if(this_validchar=="%")
							{
								valid=true;
								break;	
							}
							else if(this_validchar=="*"&&this_putchar=="")
							{
								valid=false;
								word="Value of "+action_name+" is invalid!\n"+action_name+"的输入值错误！";
								object=eval("action_value"+index+"_"+subindex);								
								break;	
							}
						}
						if(m<putvalue.length&&this_validchar!="%")
						{
							valid=false;
							word="Value of "+action_name+" is invalid!\n"+action_name+"的输入值错误！";
							object=eval("action_value"+index+"_"+subindex);							
						}
						if (valid==true)
						{
							break forLoop1;	
						}
					}
				}
			}
	}
	if (valid==false)//invalid
	{		
		eval("document.form1.action_boolean"+index+"_"+subindex+".value='2';");
		eval("document.all.errorinsert"+index+".innerText=word;")
		object.blur();
		object.focus();
		object.value="";
	}
	else
	{		
		eval("document.form1.action_boolean"+index+"_"+subindex+".value='1';");
		eval("document.all.errorinsert"+index+".innerText='';")
	}
}