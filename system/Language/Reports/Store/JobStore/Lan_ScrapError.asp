<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchCreateTime="Create Time|���ʱ��" 
var arrSearchCreateTime=strSearchCreateTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strSearchPlaner="Planer|�ƻ���" 
var arrSearchPlaner=strSearchPlaner.split("|")
var strSearchProgress="Progress|����" 
var arrSearchProgress=strSearchProgress.split("|")
var strSearchProgressTypeSelect="-- Select Type --|-- ѡ������ --" 
var arrSearchProgressTypeSelect=strSearchProgressTypeSelect.split("|")
var strSearchProgressFinished="Finished|������" 
var arrSearchProgressFinished=strSearchProgressFinished.split("|")
var strSearchProgressProgressing="Progressing|������" 
var arrSearchProgressProgressing=strSearchProgressProgressing.split("|")

var strBrowse="Browse Job Store Conflict with ERP |�������ͻ����ERP�Ƚϣ�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strPlaner="Planer|�ƻ���" 
var arrPlaner=strPlaner.split("|")
var strStoreStatus="Store Status|���״̬" 
var arrStoreStatus=strStoreStatus.split("|")
var strStartError="Start Error|�ƻ�������ͬ" 
var arrStartError=strStartError.split("|")
var strExceedError="Exceed Error|�������" 
var arrExceedError=strExceedError.split("|")
var strStartQuantity="Start Quantity|�ƻ�����" 
var arrStartQuantity=strStartQuantity.split("|")
var strCompleteQuantity="Complete Quantity|�������" 
var arrCompleteQuantity=strCompleteQuantity.split("|")
var strScrapQuantity="Scrap Quantity|��������" 
var arrScrapQuantity=strScrapQuantity.split("|")
var strERPStartQuantity="ERP Start Quantity|ERP �ƻ�����" 
var arrERPStartQuantity=strERPStartQuantity.split("|")
var strERPCompleteQuantity="ERP Complete Quantity|ERP �������" 
var arrERPCompleteQuantity=strERPCompleteQuantity.split("|")
var strERPScrapQuantity="ERP Scrap Quantity|ERP ��������" 
var arrERPScrapQuantity=strERPScrapQuantity.split("|")
var strERPUpdateTime="Update Time|����ʱ��" 
var arrERPUpdateTime=strERPUpdateTime.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchCreateTime.innerText=arrSearchCreateTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchPlaner.innerText=arrSearchPlaner[<%=session("language")%>]}catch(e){}
try{inner_SearchProgress.innerText=arrSearchProgress[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[0].text=arrSearchProgressTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[1].text=arrSearchProgressFinished[<%=session("language")%>]}catch(e){}
try{document.all.progress.options[2].text=arrSearchProgressProgressing[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Planer.innerText=arrPlaner[<%=session("language")%>]}catch(e){}
try{inner_StoreStatus.innerText=arrStoreStatus[<%=session("language")%>]}catch(e){}
try{inner_StartError.innerText=arrStartError[<%=session("language")%>]}catch(e){}
try{inner_ExceedError.innerText=arrExceedError[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_CompleteQuantity.innerText=arrCompleteQuantity[<%=session("language")%>]}catch(e){}
try{inner_ScrapQuantity.innerText=arrScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_ERPStartQuantity.innerText=arrERPStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_ERPCompleteQuantity.innerText=arrERPCompleteQuantity[<%=session("language")%>]}catch(e){}
try{inner_ERPScrapQuantity.innerText=arrERPScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_ERPUpdateTime.innerText=arrERPUpdateTime[<%=session("language")%>]}catch(e){}
}
</script>