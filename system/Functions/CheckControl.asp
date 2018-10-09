<script language="javascript">
function checkall()
{
	for(var i=1;i<<%=i%>;i++)
	{
		if (eval("document.checkform.id"+i+".disabled==false"))
		{
		eval("document.checkform.id"+i+".checked=true;")
		}
	}
}
function uncheckall()
{	
	for(var i=1;i<<%=i%>;i++)
	{
	eval("document.checkform.id"+i+".checked=false;")
	}
}

<%if i<>"" and j<>"" then%>
function checkalla()
{
	for(var i=1;i<<%=i%>;i++)
	{
		eval("document.checkform.nid"+i+".checked=true;")
	}
	for(var j=1;j<<%=j%>;j++)
	{
		eval("document.checkform.pid"+j+".checked=true;")
	}
}

function uncheckalla()
{	
	for(var i=1;i<<%=i%>;i++)
	{
		eval("document.checkform.nid"+i+".checked=false;")
	}
	for(var j=1;j<<%=j%>;j++)
	{
		eval("document.checkform.pid"+j+".checked=false;")
	}
}

function checkallc(condition,type)
{
	if (type=="station"||type=="all")
	{
		for(var i=1;i<<%=i%>;i++)
		{
			var value=eval("document.checkform.nfactory"+i+".value")
			if (value==condition)
			{
			eval("document.checkform.nid"+i+".checked=true;")
			}
			else
			{
			eval("document.checkform.nid"+i+".checked=false;")
			}
		}
	}
	if (type=="part"||type=="all")
	{
		for(var j=1;j<<%=j%>;j++)
		{	
			var value=eval("document.checkform.pfactory"+j+".value")
			if (value==condition)
			{
			eval("document.checkform.pid"+j+".checked=true;")
			}
			else
			{
			eval("document.checkform.pid"+j+".checked=false;")
			}
		}
	}
}

function uncheckallc(condition,type)
{	
	if (type=="station"||type=="all")
	{
		for(var i=1;i<<%=i%>;i++)
		{
			var value=eval("document.checkform.nfactory"+i+".value")
			if (value==condition)
			{
			eval("document.checkform.nid"+i+".checked=false;")
			}
		}
	}
	if (type=="part"||type=="all")
	{
		for(var j=1;j<<%=j%>;j++)
		{
			var value=eval("document.checkform.pfactory"+j+".value")
			if (value==condition)
			{
			eval("document.checkform.pid"+j+".checked=false;")
			}
		}
	}
}
<%end if%>

function duplicaterow(i,q)
{
	with(document.checkform)
	{
		if (eval("tracking"+q+".value!=''&&tracking"+i+".value==''"))
		{
		var PI=new Number(eval("PT_"+eval("tracking_code"+i+".value")+".value"))
		var TN="WW"+weekindex.value+eval("tracking_code"+i+".value")+"-"+PI
		eval("tracking"+i+".value='"+TN+"'")
		eval("PT_"+eval("tracking_code"+i+".value")+".value='"+(PI+1)+"'")
		}
	}
}
function duplicateup(o)
{	
//	var prodoc=open('/Components/Progress.html','','width=300,height=50,scrollbars=yes,resizable=yes');
//	prodoc.moveTo(screen.width/2,screen.height/2);
	for (var p=2;p<=o;p++)
	{
//	prodoc.document.write(Math.round(p/o*250));
//	prodoc.document.all.Proper.innerText=Math.round(p/o*100)+"%";
	duplicaterow(p,p-1);
	}
//	prodoc.close();
//	prodoc=null;
}
function generatetracking()
{
	with(document.checkform)
	{
		var TN="WW"+weekindex.value+tracking_code1.value+"-"+1
		tracking1.value=TN
		eval("PT_"+tracking_code1.value+".value=2")
	}
}
function layopen()
{document.all.JProgress.style.visibility="visible";}
function layclose()
{document.all.JProgress.style.visibility="hidden";}
</script>