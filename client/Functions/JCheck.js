// JavaScript Document
function numbercheck(ob,errorinsert)
{
	if(isNumberString(ob.value,'1234567890.')!=1)
	{
	errorinsert.innerText="Number format is error!\n数字格式错误！";
	ob.value="";
	ob.blur();
	ob.focus();
	}
	else
	{errorinsert.innerText="";}
}
	
function numbercheck2(ob,i)
{
	eval("var ei=document.all.errorinsert"+i)
	if(isNumberString(ob.value,'1234567890.')!=1)
	{
	ei.innerText="Number format is error!\n数字格式错误！";
	ob.value="";
	ob.blur();
	ob.focus();
	}
	else
	{
	ei.innerText="";
	}	
}

function numbercheck3(ob,i)
{
	eval("var ei=document.all.derrorinsert"+i)
	if(isNumberString(ob.value,'1234567890')!=1)
	{
	ei.innerText="Number format is error!\n数字格式错误！";
	ob.value=0;
	ob.blur();
	ob.focus();
	}
	else
	{
	ei.innerText="";
	}	
}


function isNumberString (InString,RefString)
{
if(InString.length==0) return (false);
	for (Count=0; Count < InString.length; Count++)
	{
		TempChar= InString.substring (Count, Count+1);
		if (RefString.indexOf (TempChar, 0)==-1)  
		return (false);
	}
return (true);
}

function submitonce(theform) {
    if (document.all || document.getElementById) {
        var formObj = eval("document." + theform);        
        for (i = 0; i < formObj.length; i++) {
            var tempobj = formObj.elements[i]
            if (tempobj.type.toLowerCase() == "submit" || tempobj.type.toLowerCase() == "reset") {
                tempobj.disabled = !tempobj.disabled                
            }
        }
    }    
}
String.prototype.trim=function(){
   return this.replace(/(^\s*)|(\s*$)/g, "");
}
String.prototype.ltrim=function(){
   return this.replace(/(^\s*)/g,"");
}
String.prototype.rtrim=function(){
   return this.replace(/(\s*$)/g,"");
}