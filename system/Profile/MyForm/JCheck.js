// JavaScript Document
function getParam()
{
	with(document.form1)
	{
		if(thisform.selectedIndex!=0)
		{
			document.all.paramFrame.src="/Functions/getFormParam.asp?id="+thisform.options[thisform.selectedIndex].value;
		}
		else
		{
			document.all.paramhtml1.innerHTML="";
			document.all.paramhtml2.innerHTML="";
			document.all.paramhtml3.innerHTML="";
			document.all.jobinfohtml.innerHTML="";
			document.all.approveflow.innerHTML="";
			document.all.actperson.innerHTML="";
		}
	}
}
function getJobInfo()
{
	with(document.form1)
	{
		if(param1.value!="")
		{
			document.all.JobInfoFrame.src="/Functions/getFormJobInfo.asp?jobnumber="+param1.value;
		}
		else
		{
			document.all.jobinfohtml.innerHTML="";
		}
	}
}