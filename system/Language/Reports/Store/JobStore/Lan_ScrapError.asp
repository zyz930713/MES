<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchCreateTime="Create Time|入库时间" 
var arrSearchCreateTime=strSearchCreateTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchPlaner="Planer|计划人" 
var arrSearchPlaner=strSearchPlaner.split("|")
var strSearchProgress="Progress|进度" 
var arrSearchProgress=strSearchProgress.split("|")
var strSearchProgressTypeSelect="-- Select Type --|-- 选择类型 --" 
var arrSearchProgressTypeSelect=strSearchProgressTypeSelect.split("|")
var strSearchProgressFinished="Finished|入库完成" 
var arrSearchProgressFinished=strSearchProgressFinished.split("|")
var strSearchProgressProgressing="Progressing|进行中" 
var arrSearchProgressProgressing=strSearchProgressProgressing.split("|")

var strBrowse="Browse Job Store Conflict with ERP |浏览入库冲突（与ERP比较）" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strPlaner="Planer|计划人" 
var arrPlaner=strPlaner.split("|")
var strStoreStatus="Store Status|入库状态" 
var arrStoreStatus=strStoreStatus.split("|")
var strStartError="Start Error|计划数量不同" 
var arrStartError=strStartError.split("|")
var strExceedError="Exceed Error|数量溢出" 
var arrExceedError=strExceedError.split("|")
var strStartQuantity="Start Quantity|计划数量" 
var arrStartQuantity=strStartQuantity.split("|")
var strCompleteQuantity="Complete Quantity|完成数量" 
var arrCompleteQuantity=strCompleteQuantity.split("|")
var strScrapQuantity="Scrap Quantity|报废数量" 
var arrScrapQuantity=strScrapQuantity.split("|")
var strERPStartQuantity="ERP Start Quantity|ERP 计划数量" 
var arrERPStartQuantity=strERPStartQuantity.split("|")
var strERPCompleteQuantity="ERP Complete Quantity|ERP 完成数量" 
var arrERPCompleteQuantity=strERPCompleteQuantity.split("|")
var strERPScrapQuantity="ERP Scrap Quantity|ERP 报废数量" 
var arrERPScrapQuantity=strERPScrapQuantity.split("|")
var strERPUpdateTime="Update Time|更新时间" 
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