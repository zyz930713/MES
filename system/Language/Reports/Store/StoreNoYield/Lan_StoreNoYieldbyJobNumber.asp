<script language="javascript">
var strBrowse="Check No Yield (sort by job number)|���鲻��������� (������������)" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strSortTitle="Sort By|����ʽ" 
var arrSortTitle=strSortTitle.split("|")
var strSortPartNumber="Part Number|�ͺ�" 
var arrSortPartNumber=strSortPartNumber.split("|")
var strSortNormal="Normal|����" 
var arrSortNormal=strSortNormal.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strNoYieldCheck="No Yield Check|�������ʸ���" 
var arrNoYieldCheck=strNoYieldCheck.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strInputQuantity="Input Quantity|��������" 
var arrInputQuantity=strInputQuantity.split("|")
var strStoreQuantity="Store Quantity|�������" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strYield="Yield|��ʱ����" 
var arrYield=strYield.split("|")
var strStoreTime="Store Time|���ʱ��" 
var arrStoreTime=strStoreTime.split("|")
var strStoreType="Store Type|�������" 
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