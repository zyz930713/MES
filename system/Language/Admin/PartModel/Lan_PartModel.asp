<script language="javascript">
var strSearch="Search Model|查询型号"
var arrSearch=strSearch.split("|")
var strSearchModelNumber="Model Number|型号" 
var arrSearchModelNumber=strSearchModelNumber.split("|")
var strSearchFactory="Factory|所属厂" 
var arrSearchFactory=strSearchFactory.split("|")
var strBrowse="Browse Model|浏览型号" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strModelNumber="Model Number|型号" 
var arrModelNumber=strModelNumber.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strDescription="Description|描述" 
var arrDescription=strDescription.split("|")
var strWIPInventory="WIP Inventory|库位" 
var arrWIPInventory=strWIPInventory.split("|")
var strLeadTime="Lead Time|完成时间" 
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
