<script language="javascript">
var strSearch="Search Store Records|搜索入库记录" 
var arrSearch=strSearch.split("|")
var strSearchtoreTime="Change Time|修改时间" 
var arrSearchStoreTime=strSearchtoreTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Store Change Records|浏览入库修改记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strInputQuantity="Input Quantity|生产数量" 
var arrInputQuantity=strInputQuantity.split("|")
var strStoreQuantity="Store Quantity|入库数量" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strStoreTime="Store Time|入库时间" 
var arrStoreTime=strStoreTime.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchStoreTime.innerText=arrSearchStoreTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_StoreQuantity.innerText=arrStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_StoreTime.innerText=arrStoreTime[<%=session("language")%>]}catch(e){}
}
</script>