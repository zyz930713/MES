// JavaScript Document
function getParam(jobnumber)
{
	with(document.form1)
	{
		if(thisform.selectedIndex!=0)
		{
			document.all.paramFrame.src="/Functions/getFormParam.asp?id="+thisform.options[thisform.selectedIndex].value+"&jobnumber="+jobnumber;
		}
		else
		{
			document.all.paramhtml1.innerHTML="";
			document.all.paramhtml2.innerHTML="";
			document.all.paramhtml3.innerHTML="";
			document.all.paramhtml4.innerHTML="";
			document.all.jobinfohtml.innerHTML="";
			document.all.approveflow.innerHTML="";
			document.all.actperson.innerHTML="";
		}
	}
}
function getEditParam()
{
	with(document.form1)
	{
		if(formid.value!="")
		{
			document.all.paramFrame.src="/Functions/getFormParam.asp?id="+formid.value+"&paramvalue1="+paramvalue1.value+"&paramvalue2="+paramvalue2.value+"&paramvalue3="+paramvalue3.value;
		}
		else
		{
			document.all.paramhtml1.innerHTML="";
			document.all.paramhtml2.innerHTML="";
			document.all.paramhtml3.innerHTML="";
			document.all.paramhtml4.innerHTML="";
			document.all.jobinfohtml.innerHTML="";
			document.all.approveflow.innerHTML="";
			document.all.actperson.innerHTML="";
		}
	}
}
function getJobInfo(type,point)
{
	//type=0 'job quantity decrease
	//type=1 'job line changed
	//type=2 'job part number changed
	//type=3 'job cancelled
	with(document.form1)
	{
		if(param1.value!="")
		{
			if (type==0)
			{
			document.all.JobInfoFrame.src="/Functions/getFormJobInfo.asp?jobnumber="+param1.value;
			}
			else
			{
			eval("var objectname=paramname"+point);
			eval("var objectindex=paramindex"+point);
			document.all.JobInfoFrame.src="/Functions/getFormJobInfo.asp?jobnumber="+param1.value+"&objectname="+objectname.value+"&objectindex="+objectindex.value;
			}
		}
		else
		{
			document.all.jobinfohtml.innerHTML="";
		}
	}
}
function compareJobQuantity()
{
	with(document.form1)	
	{
		if(isNumberString(param2.value,'1234567890')!=1)
		{
			alert("Format of new Quanity is not correct!\n新的数量格式错误！");
			param2.value="";
			return false;
		}
		else
		{
			if(new Number(param2.value)>=new Number(start_quantity.value))
			{
				alert("New Quanity is equal or exceed original quantity!\n新的数量大于或等于原来的数量！")	;
				alert("Please click job info button to try!\n点击[工单信息]按钮，再试一下！")	;
				//param2.value="";
				return false;
			}
		}
	}
}