// JavaScript Document
function item_move(fbox,tbox)
{
/*	for(var i=0; i<fbox.options.length; i++)
	{
		if(fbox.options[i].selected && fbox.options[i].value != "") 
		{
		var no = new Option();
		no.value = fbox.options[i].value;
		no.text = fbox.options[i].text;
		tbox.options[tbox.options.length] = no;
		fbox.options[i]=null;
		}
	}*/
	var i=0;
	var j=fbox.options.length-1;
	do
	{
		if(fbox.options[i].selected && fbox.options[i].value != "") 
		{
		var no = new Option();
		no.value = fbox.options[i].value;
		no.text = fbox.options[i].text;
		tbox.options[tbox.options.length] = no;
		fbox.options[i]=null;
		j--;
		}
		else
		{
		i++
		};
	}
	while(i<=j);
}

function item_all(fbox,tbox)
{	
	k=fbox.options.length;
	for(var i=0;i<k;i++)
	{
		fbox.options[i].selected=false;
	}
	for(var i=0;i<k;i++)
	{
		fbox.options[0].selected=true;
		item_move(fbox,tbox);
	}
}

function item_up(box)
{
	var k=null;
	for(var i=0; i<box.options.length; i++)
	{
		if (box.options[i].selected&&box.options[i]!=""&&box.selectedIndex!=0)
		{
		var tmpval = box.options[i].value;
		var tmpval2 = box.options[i].text;
		box.options[i].value = box.options[i-1].value;
		box.options[i].text = box.options[i-1].text
		box.options[i-1].value = tmpval;
		box.options[i-1].text = tmpval2;
		box.options[i].selected=false;
		box.options[i-1].selected=false;
		var k=i-1;
		}
	}
	if (k!=null)
	{
		box.options[k].selected=true;
	}
}

function item_down(box)
{
	var k=null;
	for(var i=0; i<box.options.length; i++)
	{
		if(box.options[i].selected&&box.options[i] != ""&&box.selectedIndex+1!=box.options.length)
		{
		var tmpval = box.options[i].value;
		var tmpval2 = box.options[i].text;
		box.options[i].value = box.options[i+1].value;
		box.options[i].text = box.options[i+1].text
		box.options[i+1].value = tmpval;
		box.options[i+1].text = tmpval2;
		box.options[i].selected=false;
		box.options[i+1].selected=false;
		var k=i+1;
		}
	}
	if (k!=null)
	{
		box.options[k].selected=true;
	}
}

function item_top(box)
{
	if(box.selectedIndex!=-1)
	{
		var tmpval = box.options[box.selectedIndex].value;
		var tmpval2 = box.options[box.selectedIndex].text;
		for(var i=box.selectedIndex;i>0;i--)
		{
			box.options[i].value = box.options[i-1].value;
			box.options[i].text = box.options[i-1].text;
		}
		box.options[0].value = tmpval;
		box.options[0].text = tmpval2;
		box.options[0].selected=true;
	}
}

function item_bottom(box)
{	
	if(box.selectedIndex!=-1)
	{
		var tmpval = box.options[box.selectedIndex].value;
		var tmpval2 = box.options[box.selectedIndex].text;
		for(var i=box.selectedIndex;i<box.options.length-1;i++)
		{
			box.options[i].value = box.options[i+1].value;
			box.options[i].text = box.options[i+1].text
		}
		box.options[box.options.length-1].value = tmpval
		box.options[box.options.length-1].text = tmpval2
		box.options[box.options.length-1].selected=true;
	}
}

function fromitem_all()
{
	for(var i=0;i<document.form1.fromitem.options.length;i++)
	{
	document.form1.fromitem.options[i].selected=true;
	}
}

function toitem_all()
{
	for(var i=0;i<document.form1.toitem.options.length;i++)
	{
	document.form1.toitem.options[i].selected=true;
	}
}

function new_toitem_all(item_name)
{
	for(var i=0;i<item_name.options.length;i++)
	{
	item_name.options[i].selected=true;
	}
}

function toitem1_all()
{
	for(var i=0;i<document.form1.toitem1.options.length;i++)
	{
	document.form1.toitem1.options[i].selected=true;
	}
}

function toitem2_all()
{
	for(var i=0;i<document.form1.toitem2.options.length;i++)
	{
	document.form1.toitem2.options[i].selected=true;
	}
}

function toitem3_all()
{
	for(var i=0;i<document.form1.toitem3.options.length;i++)
	{
	document.form1.toitem3.options[i].selected=true;
	}
}

function selectedcount()
{
	document.all.selectedinsert.innerText="("+document.form1.toitem.length+")";
}

function selectedcount1()
{
	document.all.selectedinsert1.innerText="("+document.form1.toitem1.length+")";
}

function selectedcount2()
{
	document.all.selectedinsert2.innerText="("+document.form1.toitem2.length+")";
}

function selectedcount200()
{
	document.all.selectedinsert200.innerText="("+document.form1.toitem200.length+")";
}

function selectedcount3()
{
	document.all.selectedinsert3.innerText="("+document.form1.toitem3.length+")";
}

function deselectedcount()
{
	document.all.deselectedinsert.innerText="("+document.form1.fromitem.length+")";
}

function deselectedcount1()
{
	document.all.deselectedinsert1.innerText="("+document.form1.fromitem1.length+")";
}

function deselectedcount2()
{
	document.all.deselectedinsert2.innerText="("+document.form1.fromitem2.length+")";
}

function deselectedcount3()
{
	document.all.deselectedinsert3.innerText="("+document.form1.fromitem3.length+")";
}

function parentdeselectedcount()
{
	parent.deselectedinsert.innerText="("+parent.form1.fromitem.length+")";
}