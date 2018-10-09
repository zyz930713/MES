// JavaScript Document
function isNumberString (InString,RefString)
{
if(InString.length==0) return (false);
	for (Count=0; Count < InString.length; Count++)
	{
		TempChar= InString.substring (Count, Count+1);
		if (RefString.indexOf (TempChar, 0)==-1)  
		return (false);
	}
return (true);
}
function weeknumbercheck(ob)
{
	if(isNumberString(ob.value,"1234567890")!=1||new Number(ob.value)>52||ob.value=="0")
	{
	alert("Week Number format is error!");
	ob.value="";
	}
}
function yearnumbercheck(ob)
{
	if(isNumberString(ob.value,"1234567890")!=1)
	{
	alert("Year Number format is error!");
	ob.value="";
	}
}
function monthnumbercheck(ob)
{
	if(isNumberString(ob.value,"1234567890")!=1||new Number(ob.value)>12||ob.value=="0")
	{
	alert("Month Number format is error!");
	ob.value="";
	}
}

function numbercheck(ob)
{
	if (isNumberString(ob.value,"1234567890.")!=1)
	{
	alert("Number format is error!")
	ob.value="";
	}
}

function getXMLResponse(url,pause,jfunction)
{
	if(window.ActiveXObject)
	{
		http_request=new ActiveXObject("Microsoft.XMLHTTP");
		if (http_request)
		{
		http_request.onreadystatechange = jfunction;
		http_request.open('GET',url,pause);
		http_request.send(null);
		}
		else
		{
			alert("XMLHTTP create error!");
		}
	}
	else
	{
		alert("XMLHTTP object error!");
	}
}

function hournumbercheck(ob)
{
	if(isNumberString(ob.value,"1234567890")!=1||new Number(ob.value)>24)
	{
	alert("Hour Number format is error!");
	ob.value="0";
	}
	else
	{
		if(new Number(ob.value)==24)	
		{
		ob.value="0";
		}
	}
}

function minutenumbercheck(ob)
{
	if(isNumberString(ob.value,"1234567890")!=1||new Number(ob.value)>60)
	{
	alert("Minute Number format is error!");
	ob.value="0";
	}
	else
	{
		if(new Number(ob.value)==60)	
		{
		ob.value="0";
		}
	}
}

window.onbeforeunload=function (){ 
	if (document.all || document.getElementById) {
		try{
			var formObj = eval("document.form1");        
			for (i = 0; i < formObj.length; i++) {
				var tempobj = formObj.elements[i]
				if (tempobj.type.toLowerCase() == "submit" || tempobj.type.toLowerCase() == "reset") {
					tempobj.disabled = !tempobj.disabled                
				}
			}
		}catch(e){}
    }
}
