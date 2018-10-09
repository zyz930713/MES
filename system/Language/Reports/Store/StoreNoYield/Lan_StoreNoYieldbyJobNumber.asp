<script language="javascript">
var strBrowse="Check No Yield (sort by job number)|复查不计良率入库 (按工单号排序)" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strSortTitle="Sort By|排序方式" 
var arrSortTitle=strSortTitle.split("|")
var strSortPartNumber="Part Number|型号" 
var arrSortPartNumber=strSortPartNumber.split("|")
var strSortNormal="Normal|正常" 
var arrSortNormal=strSortNormal.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strNoYieldCheck="No Yield Check|不计良率复查" 
var arrNoYieldCheck=strNoYieldCheck.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strInputQuantity="Input Quantity|生产数量" 
var arrInputQuantity=strInputQuantity.split("|")
var strStoreQuantity="Store Quantity|入库数量" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strYield="Yield|即时良率" 
var arrYield=strYield.split("|")
var strStoreTime="Store Time|入库时间" 
var arrStoreTime=strStoreTime.split("|")
var strStoreType="Store Type|入库类型" 
var arrStoreType=strStoreType.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_SortTitle.innerText=arrSortTitle[<%=session("language")%>]}catch(e){}
try{inner_SortPartNumber.innerText=arrSortPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SortNormal.innerText=arrSortNormal[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_NoYieldCheck.innerText=arrNoYieldCheck[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_StoreQuantity.innerText=arrStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_Yield.innerText=arrYield[<%=session("language")%>]}catch(e){}
try{inner_StoreTime.innerText=arrStoreTime[<%=session("language")%>]}catch(e){}
try{inner_StoreType.innerText=arrStoreType[<%=session("language")%>]}catch(e){}
}
</script>