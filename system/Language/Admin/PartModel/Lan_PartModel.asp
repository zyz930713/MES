<script language="javascript">
var strSearch="Search Model|��ѯ�ͺ�"
var arrSearch=strSearch.split("|")
var strSearchModelNumber="Model Number|�ͺ�" 
var arrSearchModelNumber=strSearchModelNumber.split("|")
var strSearchFactory="Factory|������" 
var arrSearchFactory=strSearchFactory.split("|")
var strBrowse="Browse Model|����ͺ�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strModelNumber="Model Number|�ͺ�" 
var arrModelNumber=strModelNumber.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strDescription="Description|����" 
var arrDescription=strDescription.split("|")
var strWIPInventory="WIP Inventory|��λ" 
var arrWIPInventory=strWIPInventory.split("|")
var strLeadTime="Lead Time|���ʱ��" 
var arrLeadTime=strLeadTime.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchModelNumber.innerText=arrSearchModelNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_ModelNumber.innerText=arrModelNumber[<%=session("language")%>]}catch(e){}
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Description.innerText=arrDescription[<%=session("language")%>]}catch(e){}
try{inner_WIPInventory.innerText=arrWIPInventory[<%=session("language")%>]}catch(e){}
try{inner_LeadTime.innerText=arrLeadTime[<%=session("language")%>]}catch(e){}
}
language()
</script>
