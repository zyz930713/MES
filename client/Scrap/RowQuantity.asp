<script language="JavaScript" type="text/javascript">
var j=1;
var gather_quantity=0;
gather_quantity=new Number(document.all.scrap_quantity1.value);
	for (var i=2;i<=<%=i-1%>;i++)
	{
		if(eval("document.all.job_number"+i+".value==document.all.job_number"+(i-1)+".value"))
		{
			gather_quantity=gather_quantity+new Number(eval("document.all.scrap_quantity"+i+".value"));
			j=j+1;
			eval("document.all.gather"+(i-1)+".innerText=' '");
		}
		else
		{
			eval("document.all.gather"+(i-1)+".innerText="+gather_quantity);
			j=1;
			gather_quantity=new Number(eval("document.all.scrap_quantity"+i+".value"));
		}
	}
	eval("document.all.gather"+(i-1)+".innerText="+gather_quantity);
</script>