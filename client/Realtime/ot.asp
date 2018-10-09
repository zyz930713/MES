<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>RealTime</title>
<script language="javascript">
var hkey_root,hkey_path,hkey_key
hkey_root="hkey_current_user"
hkey_path="\\software\\microsoft\\internet explorer\\pagesetup\\"
//设置网页打印的页眉页脚为空
function pagesetup_null(){
try{
var regwsh = new activexobject("wscript.shell")
hkey_key="header" 
regwsh.regwrite(hkey_root+hkey_path+hkey_key,"")
hkey_key="footer"
regwsh.regwrite(hkey_root+hkey_path+hkey_key,"")
}catch(e){}
}
//设置网页打印的页眉页脚为默认值
function pagesetup_default(){
try{
var regwsh = new activexobject("wscript.shell")
hkey_key="header" 
regwsh.regwrite(hkey_root+hkey_path+hkey_key,"&w&b页码，&p/&p")
hkey_key="footer"
regwsh.regwrite(hkey_root+hkey_path+hkey_key,"&u&b&d")
}catch(e){}
}
</script>
</head> 

<body><br/><br/><br/><br/><br/><br/><p align=center>
<input type="button" value="清空页码" onclick=pagesetup_null()>
<input type="button" value="恢复页码" onclick=pagesetup_default()><br/>
</p></body></html>


