// JavaScript Document
function formcheckSave()
{
	with(document.form1)
	{
		if(isWW.checked)
		{
			if(wwnumber.value=="")
			{
				alert("Week Number cannot be blank!");
				return false;
			}
		}
		if(finalfamily_name.value=="")
		{
			alert("Report Name cannot be blank!");
			return false;
		}
	}
}

function formcheckUpdate()
{
	with(document.form1)
	{
		if(isWW.checked)
		{
			if(wwnumber.value=="")
			{
				alert("Week Number cannot be blank!");
				return false;
			}
		}
	}
}