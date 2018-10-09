function formcheck()
{
	with(document.form1)
	{
		if(report_name.value=="")
		{
		alert("Report Name cannot be blank");
		return false;
		}
		if(fromdate.value=="")
		{
		alert("Report Name cannot be blank");
		return false;
		}
		if(todate.value=="")
		{
		alert("Report Name cannot be blank");
		return false;
		}
		if(factory.selectedIndex==0)
		{
		alert("Factory cannot be blank");
		return false;
		}
		if(report_type.selectedIndex==0)
		{
		alert("Report Type cannot be blank");
		return false;
		}
		if(this_station.selectedIndex==0)
		{
		alert("Station (Numerator) cannot be blank");
		return false;
		}
		if(this_action.selectedIndex==0)
		{
		alert("Action (Numerator) cannot be blank");
		return false;
		}
		if(refer_station.selectedIndex==0)
		{
		alert("Station (Denominator) cannot be blank");
		return false;
		}
		if(refer_action.selectedIndex==0)
		{
		alert("Action (Denominator) cannot be blank");
		return false;
		}
		if (toitem.length!=0)
		{
		toitem_all();
		}
	}
}

function ChangeFactory()
{
	with(document.form1)
	{
		if(factory.selectedIndex!=0)
		{
		var jfunction=queryStation;
		var url="/Functions/GetStationOfFactory.asp?factory="+factory.options[factory.options.selectedIndex].value
		getXMLResponse(url,true,jfunction);
		}
	}
}

function queryStation()
{
	with(document.form1)
	{
		if(http_request.readyState==4)
		{	
			if(http_request.status==200)
			{
				do{this_station.options[0]=null;}
				while(this_station.length=0);
				var newoption=document.createElement("option");
				newoption.text="--Select Station--";
				newoption.value="";
				this_station.options.add(newoption);
				do{refer_station.options[0]=null;}
				while(refer_station.length=0);
				var newoption=document.createElement("option");
				newoption.text="--Select Station--";
				newoption.value="";
				refer_station.options.add(newoption);
				var options_value=http_request.responseText;
				var array_options_value=options_value.split("|");
				for (var i=0;i<array_options_value.length;i++)
				{
					var array_options_item=array_options_value[i].split("*");
					var newElem=document.createElement("option");
					newElem.text=array_options_item[1];
					newElem.value=array_options_item[0];
					newElem.title=array_options_item[1];
					this_station.options.add(newElem);
				}
				for (var i=0;i<array_options_value.length;i++)
				{
					var array_options_item=array_options_value[i].split("*");
					var newElem=document.createElement("option");
					newElem.text=array_options_item[1];
					newElem.value=array_options_item[0];
					newElem.title=array_options_item[1];
					refer_station.options.add(newElem);
				}
			var jfunction=queryFamily;
			var url="/Functions/GetFamilyOfFactory.asp?factory="+factory.options[factory.options.selectedIndex].value
			getXMLResponse(url,true,jfunction);
			}
		}
	}
}

function queryFamily()
{
	with(document.form1)
	{
		if(http_request.readyState==4)
		{	
			if(http_request.status==200)
			{
				do{family.options[0]=null;}
				while(family.length=0);
				var newoption=document.createElement("option");
				newoption.text="--Select Family--";
				newoption.value="";
				family.options.add(newoption);
				var options_value=http_request.responseText;
				var array_options_value=options_value.split("|");
				for (var i=0;i<array_options_value.length;i++)
				{
					var array_options_item=array_options_value[i].split("*");
					var newElem=document.createElement("option");
					newElem.text=array_options_item[1];
					newElem.value=array_options_item[0];
					newElem.title=array_options_item[1];
					family.options.add(newElem);
				}
			}
		}
	}
}

function ChangeStation()
{
	with(document.form1)
	{
		if(factory.selectedIndex!=0)
		{
		var jfunction=queryAction;
		var url="/Functions/GetActionOfStation.asp?station="+this_station.options[this_station.options.selectedIndex].value;
		getXMLResponse(url,true,jfunction);
		}
	}
}

function queryAction()
{
	with(document.form1)
	{
		if(http_request.readyState==4)
		{	
			if(http_request.status==200)
			{
				do{this_action.options[0]=null;}
				while(this_action.length=0);
				var newoption=document.createElement("option");
				newoption.text="--Select Station--";
				newoption.value="";
				this_action.options.add(newoption);
				var options_value=http_request.responseText;
				var array_options_value=options_value.split("|");
				for (var i=0;i<array_options_value.length;i++)
				{
					var array_options_item=array_options_value[i].split("*");
					var newElem=document.createElement("option");
					newElem.text=array_options_item[1];
					newElem.value=array_options_item[0];
					newElem.title=array_options_item[1];
					this_action.options.add(newElem);
				}
			}
		}
	}
}

function ChangeReferStation()
{
	with(document.form1)
	{
		if(refer_station.selectedIndex!=0)
		{
		refer_station_name.value=refer_station.options[refer_station.options.selectedIndex].title;
		var jfunction=queryReferAction;
		var url="/Functions/GetActionOfStation.asp?station="+refer_station.options[refer_station.options.selectedIndex].value;
		getXMLResponse(url,true,jfunction);
		}
	}
}

function queryReferAction()
{
	with(document.form1)
	{
		if(http_request.readyState==4)
		{	
			if(http_request.status==200)
			{
				do{refer_action.options[0]=null;}
				while(refer_action.length=0);
				var newoption=document.createElement("option");
				newoption.text="--Select Station--";
				newoption.value="";
				refer_action.options.add(newoption);
				var options_value=http_request.responseText;
				var array_options_value=options_value.split("|");
				for (var i=0;i<array_options_value.length;i++)
				{
					var array_options_item=array_options_value[i].split("*");
					var newElem=document.createElement("option");
					newElem.text=array_options_item[1];
					newElem.value=array_options_item[0];
					newElem.title=array_options_item[1];
					refer_action.options.add(newElem);
				}
			}
		}
	}
}

function  ChangeThisAction()
{
	with(document.form1)
	{
		if (this_action.options.selectedIndex!=0)
		{
			this_action_name.value=this_action.options[this_action.options.selectedIndex].title;
		}
	}
}

function  ChangeReferAction()
{
	with(document.form1)
	{
		if (refer_action.options.selectedIndex!=0)
		{
			refer_action_name.value=refer_action.options[refer_action.options.selectedIndex].title;
		}
	}
}