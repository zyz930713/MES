// JavaScript Document
function formcheck()
{
	with(document.form1)
	{
		if (groupname.value=="")
		{
		alert("Name cannot be blank!\n���Ʋ���Ϊ�գ�");
		return false;
		}
		if (groupchinesename.value=="")
		{
		alert("Chinese Name cannot be blank!\n�������Ʋ���Ϊ�գ�");
		return false;
		}
		if (factory.selectedIndex==0)
		{
		alert("Factory cannot be blank!\n��������Ϊ�գ�")
		return false;
		}
		if (grouptype.selectedIndex==0)
		{
		alert("Group Type cannot be blank!\nȺ�������Ϊ�գ�")
		return false;
		}
	}
}