<script language="javascript">
var strWelcome="Welcome to <%=application("SystemName")%>|欢迎访问 <%=application("SystemName")%>"
var arrWelcome=strWelcome.split("|")
var strLogon ="Logon information|登陆信息" 
var arrLogon=strLogon.split("|")
var strCode="User Code|工号" 
var arrCode=strCode.split("|")
var strName="User Name|姓名" 
var arrName=strName.split("|")
var strEmail="Email|电子邮件" 
var arrEmail=strEmail.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strRole="Role|角色" 
var arrRole=strRole.split("|")
var strLinks ="Hot Linkes|常用链接" 
var arrLinks=strLinks.split("|")
var strLinksScrap ="Job Scrap|工单报废" 
var arrLinksScrap=strLinksScrap.split("|")
var strLinksStoreConfirm ="Job Store Confirm|工单入库确认" 
var arrLinksStoreConfirm=strLinksStoreConfirm.split("|")
var strLinksSubStoreConfirm ="Job Store Confirm|工单入库确认" 
var arrLinksSubStoreConfirm=strLinksSubStoreConfirm.split("|")
var strLinksScrapConfirm ="Job Scrap Confirm|工单报废确认" 
var arrLinksScrapConfirm=strLinksScrapConfirm.split("|")
function language()
{
try{inner_Welcome.innerText=arrWelcome[<%=session("language")%>]}catch(e){} 
try{inner_Logon.innerText=arrLogon[<%=session("language")%>]}catch(e){} 
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_Name.innerText=arrName[<%=session("language")%>]}catch(e){}
try{inner_Email.innerText=arrEmail[<%=session("language")%>]}catch(e){}
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Role.innerText=arrRole[<%=session("language")%>]}catch(e){}
try{inner_Links.innerText=arrLinks[<%=session("language")%>]}catch(e){}
try{inner_LinksScrap.innerText=arrLinksScrap[<%=session("language")%>]}catch(e){}
try{inner_LinksStoreConfirm.innerText=arrLinksStoreConfirm[<%=session("language")%>]}catch(e){}
try{inner_LinksSubStoreConfirm.innerText=arrLinksSubStoreConfirm[<%=session("language")%>]}catch(e){}
try{inner_LinksScrapConfirm.innerText=arrLinksScrapConfirm[<%=session("language")%>]}catch(e){}
}
language()
</script>