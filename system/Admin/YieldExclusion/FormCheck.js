// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (yield_exclusion_name.value=="")
		{
		alert("Yield Exclusion Name cannot be blank!");
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
		toitem_all();
	}
}

function refreshmodel(objectname,filterstring_notin,filterstring_in)
{
	with(document.form1)
	{
		if(filterstring_notin.value!=""||filterstring_in.value!="")
		{
			document.all.filterFrame.src="RefreshModel.asp?objectname="+objectname+"&filterstring_notin="+filterstring_notin+"&filterstring_in="+filterstring_in;
		}
	}
}
