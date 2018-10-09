function keyhandler()
{
/*var realkey=String.fromCharCode(event.keyCode);
	if (realkey=="@")//@ is submit sign
	{
		event.keyCode=null;
		document.form1.Next.click();
	}
	if (realkey=="^")//@ is submit sign
	{
		event.keyCode=null;
		window.scrollBy(0,-600);
	}
	if (realkey=="_")//@ is submit sign
	{
		event.keyCode=null;
		window.scrollBy(0,600);
	}
	if (realkey==' ')
	{
		event.keyCode=null;
	}*/
}// JavaScript Document
function tabhandler(action,there,next)
{
/*var thiskey=String.fromCharCode(event.keyCode);
	if (thiskey=='#')//# is tab sign
	{
		event.keyCode=null;
		if(action==1)
		{
			next.focus()
		}
		else
		{
			next.blur();
		}
	}
	if (thiskey=='<')
	{
		event.keyCode=null;
		there.select();
		there.value="";
	}
	if (thiskey==' ')
	{
		event.keyCode=null;
	}*/
}
function focushandler(ob)
{
	ob.style.background="#FFFF77";
}
function blurhandler(ob)
{
	ob.style.background="#F5F5F5";
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