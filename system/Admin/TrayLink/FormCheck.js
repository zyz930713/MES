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
		if (stationname.selectedIndex==0)
		{
		alert("Station Name cannot be blank!");
		return false;
		}
		if (stationchinesename.value=="")
		{
		alert("Station Chinese Name cannot be blank!");
		return false;
		}
		if (traytype.selectedIndex==0)
		{
		alert("Tray Type cannot be blank!")
		return false;
		}
		if (traysize.value=="")
		{
		alert("Tray Size cannot be blank!");
		return false;
		}	
		else
		{
			var regExp=/^\d+(\.\d+)?$/;
			if(!regExp.test(traysize.value)){				
				alert("Tray Size cannot be a number!");
				return false;
			}
		}
	}
}