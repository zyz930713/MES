// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (seriesname.value=="")
		{
		alert("Series Name cannot be blank!");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!")
		return false;
		}
		if (section.selectedIndex==0)
		{
		alert("Section cannot be blank!");
		return false;
		}
		if (line.selectedIndex==0)
		{
		alert("Line cannot be blank!");
		return false;
		}
		if (firstyield.value=="")
		{
		alert("First Yield cannot be blank!");
		return false;
		}
		if (internal_yield.value=="")
		{
		alert("Internal Yield cannot be blank!");
		return false;
		}
		if (yield.value=="")
		{
		alert("Target Yield cannot be blank!");
		return false;
		}
		toitem_all();
	}
}

function refreshmodel(objectname,filterstring_notin,filterstring_in,filter_forcedin)
{
	with(document.form1)
	{
		if(filterstring_notin.value!=""||filterstring_in.value!="")
		{	
			if(filter_forcedin.checked)
			{
			var filter_forcedin_value="1";
			}
			else
			{
			var filter_forcedin_value="0";
			}
			document.all.filterFrame.src="RefreshModel.asp?objectname="+objectname+"&filterstring_notin="+filterstring_notin+"&filterstring_in="+filterstring_in+"&filter_forcedin="+filter_forcedin_value;
			
		}
	}
}
