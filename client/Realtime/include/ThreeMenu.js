var MenuInit = function(_TmbProduct, _TmbLine, _TmbCell, defaultProduct, defaultLine, defaultCell)
{
	var TmbProduct = document.getElementById(_TmbProduct);
	var TmbLine = document.getElementById(_TmbLine);
	var TmbCell = document.getElementById(_TmbCell);
	
	function TmbSelect(Tmb, str)
	{
		for(var i=0; i<Tmb.options.length; i++)
		{
			if(Tmb.options[i].value == str)
			{
				Tmb.selectedIndex = i;
				return;
			}
		}
	}
	function TmbAddOption(Tmb, str, obj)
	{
		var option = document.createElement("OPTION");
		Tmb.options.add(option);
		option.innerHTML = str;
		option.value = str;
		option.obj = obj;
	}
	
	function changeLine()
	{
		TmbCell.options.length = 0;
		if(TmbLine.selectedIndex == -1)return;
		var item = TmbLine.options[TmbLine.selectedIndex].obj;
		for(var i=0; i<item.NextList.length; i++)
		{
			TmbAddOption(TmbCell, item.NextList[i], null);
		}
		TmbSelect(TmbCell, defaultCell);
	}
	function changeProduct()
	{
		TmbLine.options.length = 0;
		TmbLine.onchange = null;
		if(TmbProduct.selectedIndex == -1)return;
		var item = TmbProduct.options[TmbProduct.selectedIndex].obj;
		for(var i=0; i<item.ProductList.length; i++)
		{
			TmbAddOption(TmbLine, item.ProductList[i].name, item.ProductList[i]);
		}
		TmbSelect(TmbLine, defaultLine);
		changeLine();
		TmbLine.onchange = changeLine;
	}
	
	for(var i=0; i<provinceList.length; i++)
	{
		TmbAddOption(TmbProduct, provinceList[i].name, provinceList[i]);
	}
	TmbSelect(TmbProduct, defaultProduct);
	changeProduct();
	TmbProduct.onchange = changeProduct;
}

var provinceList = [
{name:'RA', ProductList:[		   
{name:'3', NextList:['1','2','3','4','5']},		   
{name:'4', NextList:['1','2','3','4','5']},
{name:'5', NextList:['1','2','3','4','5']}
]},
{name:'Donau Slim', ProductList:[		   
{name:'6', NextList:['1','2','3','4','5']}
]},
{name:'PETRA', ProductList:[		   
{name:'1', NextList:['1','2','3']},
{name:'2', NextList:['1','2','3']}
]}
];