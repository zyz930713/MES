// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (partnumber.value=="")
		{
		alert("Part Number cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!");
		return false;
		}
		if (section.selectedIndex==0)
		{
		alert("Section cannot be blank!");
		return false;
		}
		toitem1_all();	
		if (toitem2.length==0)
		{
		alert("Selected Stations cannot be blank!");
		return false;
		}
		else
		{
		toitem2_all();
		stationscount.value=toitem2.length;
			if(maxinterval.value!="")
			{
			var vmax=maxinterval.value.replace(" ","")
			var amax=vmax.split(",")
				if (amax.length!=toitem2.length-1)
				{
					alert("Format of Max Interval between Stations is wrong!");
					return false;
				}
			}
		}
	}
}

function maxintervaltip()
{
	if(document.form1.maxinterval.value!="")
	{
	var vmax=document.form1.maxinterval.value.replace(" ","")
	var amax=vmax.split(",")
	document.all.maxintervalinsert.innerText="("+(amax.length+1)+" Stations)"
	}
}