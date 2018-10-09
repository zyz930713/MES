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

function weeknumbercheck()
{
	if(isNumberString(document.form1.wwnumber.value,"1234567890")!=1||new Number(document.form1.wwnumber.value)>52||document.form1.wwnumber.value=="0")
	{
	alert("Week Number format is error!");
	document.form1.wwnumber.value="";
	}
}
function yearnumbercheck()
{
	if(isNumberString(document.form1.yearnumber.value,"1234567890")!=1)
	{
	alert("Year Number format is error!");
	document.form1.yearnumber.value="";
	}
}
function monthnumbercheck()
{
	if(isNumberString(document.form1.monthnumber.value,"1234567890")!=1||new Number(document.form1.monthnumber.value)>12||document.form1.monthnumber.value=="0")
	{
	alert("Month Number format is error!");
	document.form1.monthnumber.value="";
	}
}