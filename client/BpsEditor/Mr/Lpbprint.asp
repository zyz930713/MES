<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"> 
<HTML> 
<HEAD> 
<TITLE> New Document </TITLE> 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<script> 
var hkey_root,hkey_path,hkey_key 
hkey_root="HKEY_CURRENT_USER" 
hkey_path="file://software//Microsoft//Internet Explorer\\PageSetup\\" 
//设置网页打印的页眉页脚为空 
function pagesetup_null(){ 
try{ 
var RegWsh = new ActiveXObject("WScript.Shell") 
hkey_key="header" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"") 
hkey_key="footer" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"") 
}catch(e){} 
} 
//设置网页打印的页眉页脚为默认值 
function pagesetup_default(){ 
try{ 
var RegWsh = new ActiveXObject("WScript.Shell") 
hkey_key="header" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"&w&b页码，&p/&P") 
hkey_key="footer" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"&u&b&d") 
}catch(e){} 
} 
function setdivhidden(id){//把指定id以外的层统统隐藏 
var divs=document.getElementsByTagName("DIV"); 
for(var i=0;i<divs.length;i++) 
{ 
if(divs.item(i).id!=id) 
divs.item(i).style.display="none"; 
} 
} 
function setdivvisible(id){//把指定id以外的层统统显示 
var divs=document.getElementsByTagName("DIV"); 
for(var i=0;i<divs.length;i++) 
{ 
if(divs.item(i).id!=id) 
divs.item(i).style.display="block"; 
} 
} 
function printpr() //预览函数 
{ 
pagesetup_null();//预览之前去掉页眉，页脚 
setdivhidden("div1");//打印之前先隐藏不想打印输出的元素 
var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>'; 
document.body.insertAdjacentHTML('beforeEnd', WebBrowser);//在body标签内加入html（WebBrowser activeX控件） 
WebBrowser1.ExecWB(7, 1);//打印预览 
WebBrowser1.outerHTML = "";//从代码中清除插入的html代码 
pagesetup_default();//预览结束后页眉页脚恢复默认值 
setdivvisible("div1");//预览结束后显示按钮 
} 
function print() //打印函数 
{ 
pagesetup_null();//打印之前去掉页眉，页脚 
setdivhidden("div1"); //打印之前先隐藏不想打印输出的元素 
var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>'; 
document.body.insertAdjacentHTML('beforeEnd', WebBrowser);//在body标签内加入html（WebBrowser activeX控件） 
WebBrowser1.ExecWB(6, 1);//打印 
WebBrowser1.outerHTML = "";//从代码中清除插入的html代码 
pagesetup_default();//打印结束后页眉页脚恢复默认值 
setdivvisible("div1");//打印结束后显示按钮 
} 


function preview() 
{ 
bdhtml=window.document.body.innerHTML; 
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr)); 
window.document.body.innerHTML=prnhtml; 
window.print(); 
}

</script> 
<body> 
<div id="div0"> 
<input type="button" value="打印预览" onclick="printpr()" /> 
<input type="button" value="打印" onClick="preview()" /> 
表格一： 
</div>

<div id="div1">
	dfsafdfdfdafdafdafa
</div>
</body> 
</HTML>
