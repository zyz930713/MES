<script language="javascript">
var strWelcome="Welcome to <%=application("SystemName")%>|��ӭ���� <%=application("SystemName")%>"
var arrWelcome=strWelcome.split("|")
var strLogon ="Logon information|��½��Ϣ" 
var arrLogon=strLogon.split("|")
var strCode="User Code|����" 
var arrCode=strCode.split("|")
var strName="User Name|����" 
var arrName=strName.split("|")
var strEmail="Email|�����ʼ�" 
var arrEmail=strEmail.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strRole="Role|��ɫ" 
var arrRole=strRole.split("|")
var strLinks ="Hot Linkes|��������" 
var arrLinks=strLinks.split("|")
var strLinksScrap ="Job Scrap|��������" 
var arrLinksScrap=strLinksScrap.split("|")
var strLinksStoreConfirm ="Job Store Confirm|�������ȷ��" 
var arrLinksStoreConfirm=strLinksStoreConfirm.split("|")
var strLinksSubStoreConfirm ="Job Store Confirm|�������ȷ��" 
var arrLinksSubStoreConfirm=strLinksSubStoreConfirm.split("|")
var strLinksScrapConfirm ="Job Scrap Confirm|��������ȷ��" 
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