<script language="javascript">
var strSearch="Search Store Records|��������¼" 
var arrSearch=strSearch.split("|")
var strSearchtoreTime="Change Time|�޸�ʱ��" 
var arrSearchStoreTime=strSearchtoreTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")

var strBrowse="Browse Store Change Records|�������޸ļ�¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strInputQuantity="Input Quantity|��������" 
var arrInputQuantity=strInputQuantity.split("|")
var strStoreQuantity="Store Quantity|�������" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strStoreTime="Store Time|���ʱ��" 
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