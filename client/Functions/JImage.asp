<script language="JavaScript" type="text/javascript">
function loadimage()
{
image1=new Image(25,25);
image2=new Image(25,25);
image1.src="/Images/Radio1.gif";
image2.src="/Images/Radio2.gif";
	for (var i=1;i<=<%=i-1%>;i++)
	{
		if (eval("typeof(document.all.img"+i+")=='object'"))
		{
		eval("document.all.img"+i+".src=image1.src;")
		}
	}
}

function changeimage(j)
{
	if (eval("getRight(document.all.img"+j+".src,10)=='Radio1.gif'"))
	{
		for(var i=1;i<=<%=i-1%>;i++)
		{
			if(eval("typeof(document.all.img"+i+")=='object'"))
			{
				if (j==i)
				{
					eval("document.all.img"+i+".src=image2.src")
					document.all.selected_station.value=j;
				}
				else
				{
					eval("document.all.img"+i+".src=image1.src")
				}
			}
		}
	}
	else
	{
		document.all.selected_station.value=j;
	}
}
</script>
<script language="vbscript" type="text/vbscript">
function getRight(string,num)
getRight=right(string,num)
end function
</script>