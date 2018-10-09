// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (groupname.value=="")
		{
		alert("Name cannot be blank!\n名称不能为空！");
		return false;
		}
		if (groupchinesename.value=="")
		{
		alert("Chinese Name cannot be blank!\n中文名称不能为空！");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!\n工厂不能为空！")
		return false;
		}
		if (grouptype.selectedIndex==0)
		{
		alert("Group Type cannot be blank!\n群组类别不能为空！")
		return false;
		}
	}
}