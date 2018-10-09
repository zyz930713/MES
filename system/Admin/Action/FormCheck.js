// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (actionname.value=="")
		{
		alert("Action Name cannot be blank!");
		return false;
		}
		if (actioncode.value=="")
		{
		alert("Action Code cannot be blank!");
		return false;
		}
		
		if (actionchinesename.value=="")
		{
		alert("Action Chinese Name cannot be blank!");
		return false;
		}
		if (actionpurpose.selectedIndex==0)
		{
		alert("Action Purpose cannot be blank!")
		return false;
		}
		if (position.selectedIndex==0)
		{
		alert("Action Position cannot be blank!")
		return false;
		}
		if(!actiontype[0].checked&&!actiontype[1].checked)
		{
		alert("Action Type cannot be blank!")
		return false;
		}
		if (componenttype.selectedIndex==0)
		{
		alert("Component Type cannot be blank!")
		return false;
		}
	}
}
function formcheck_New()
{
	with(document.form1)
	{
		if (actionname.value=="")
		{
		alert("Action Name cannot be blank!");
		return false;
		}
		if (actioncode.value=="")
		{
		alert("Action Code cannot be blank!");
		return false;
		}
		
		if (actionchinesename.value=="")
		{
		alert("Action Chinese Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Action Belonged Factory cannot be blank!")
		return false;
		}
		if (actionpurpose.selectedIndex==0)
		{
		alert("Action Purpose cannot be blank!")
		return false;
		}
		if (position.selectedIndex==0)
		{
		alert("Action Position cannot be blank!")
		return false;
		}
		if(!actiontype[0].checked&&!actiontype[1].checked)
		{
		alert("Action Type cannot be blank!")
		return false;
		}
		if (componenttype.selectedIndex==0)
		{
		alert("Component Type cannot be blank!")
		return false;
		}
		
	}
}