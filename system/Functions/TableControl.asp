<script language="javascript">
<%if tabi="" then
tabi=0
end if%>
function tabshow(tabindex)
{
	for (var i=1;i<<%=tabi%>;i++)
	{
		if(i==tabindex && eval("document.all['tab"+i+"']").style.display=='none')
		{
		eval("document.all['tab"+i+"'].style.display=''");
		eval("document.all.tabimg"+i+".src='/Images/Treeimg/minus.gif'");
		//eval("document.all['headth"+i+"'].bgColor='#6699FF'");
		}
		else
		{
		eval("document.all['tab"+i+"'].style.display='none'");
		eval("document.all.tabimg"+i+".src='/Images/Treeimg/plus.gif'");
		//eval("document.all['headth"+i+"'].bgColor='#FFFFFF'");
		}
	}
}

function tableexpand(i,url,count)
{
	if(eval("document.all.tabimg"+i+".title=='Expand'"))
	{
	eval("document.all['tab"+i+"'].style.display=''");
	eval("document.all.tabimg"+i+".src='/Images/Treeimg/minus.gif'");
	eval("document.all.tabimg"+i+".title='Collapse'");
	eval("document.all.tabfrm"+i+".src='"+url+"'")
		if (count<10)
		{eval("document.all.tabfrm"+i+".height="+(count+2)*20)}
	}
	else
	{
	eval("document.all['tab"+i+"'].style.display='none'");
	eval("document.all.tabimg"+i+".src='/Images/Treeimg/plus.gif'");
	eval("document.all.tabimg"+i+".title='Expand'");
	eval("document.all.tabfrm"+i+".src=''")
	}
}

function redirect(url)
{
window.open(url,'main','')
}function redirectA(url)
{
window.open(url)
}
</script>