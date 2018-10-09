// JavaScript Document
function formcheck()
{
	var flag=0;
	with(document.form1)
	{
		if(close_fromdate.value==""&&close_todate.value=="")
		{
			alert("Must select span time!");
			return false;
		}
	}
}

function clear_close()
{
	with(document.form1)
	{
		close_fromdate.value="";
		close_fromhour.value="";
		close_fromminute.value="";
		close_todate.value="";
		close_tohour.value="";
		close_tominute.value="";
	}
}