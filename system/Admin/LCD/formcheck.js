// JavaScript Document
function typecheck(thisday,thistype,path,query)
{
	if(thistype=='1')
	{
		location.href="WelcomeChange.asp?thisday="+thisday+"&thistype=1&path="+path+"&query="+query;
	}
	else
	{
		location.href="WelcomeChange.asp?thisday="+thisday+"&thistype=2&path="+path+"&query="+query;
	}
}