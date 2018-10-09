// JavaScript Document
function formcheck()
{
	var flag=0;
	with(document.form1)
	{
		if(part_number.value==""&&job_number.value==""&&line_name.value=="")
		{
			alert("Must select at least one condition (Part Number or Job Number or Line Name)!");
			return false;
		}
		if(start_fromdate.value==""&&start_todate.value==""&&close_fromdate.value==""&&close_todate.value=="")
		{
			alert("Must select span time!");
			return false;
		}
		else
		{
			if((start_fromdate.value!=""&&start_todate.value=="")||(start_fromdate.value==""&&start_todate.value!=""))
			{
			alert("Must select full time for job start time!");
			return false;	
			}
			if((close_fromdate.value!=""&&close_todate.value=="")||(close_fromdate.value==""&&close_todate.value!=""))
			{
			alert("Must select full time for job close time!");
			return false;	
			}
		}
	}
}
function clear_start()
{
	with(document.form1)
	{
		start_fromdate.value="";
		start_fromhour.value="";
		start_fromminute.value="";
		start_todate.value="";
		start_tohour.value="";
		start_tominute.value="";
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