<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchChangeTime="Change Time|变更时间" 
var arrSearchChangeTime=strSearchChangeTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchChangeCode="Change Code|变更人" 
var arrSearchChangeCode=strSearchChangeCode.split("|")
var strSearchChangeType="Change Type|变更类型" 
var arrSearchChangeType=strSearchChangeType.split("|")
var strSearchChangeTypeSelect="-- Select Type --|-- 选择类型 --" 
var arrSearchChangeTypeSelect=strSearchChangeTypeSelect.split("|")
var strSearchChangeEdit="Edit|修改" 
var arrSearchChangeEdit=strSearchChangeEdit.split("|")
var strSearchChangeDelete="Delete|删除" 
var arrSearchChangeDelete=strSearchChangeDelete.split("|")

var strBrowse="Browse Store Records|浏览入库记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strChangeType="Change Type|变更类型" 
var arrChangeType=strChangeType.split("|")
var strChangeCode="Change Code|变更人" 
var arrChangeCode=strChangeCode.split("|")
var strChangeReason="Change Reason|变更理由" 
var arrChangeReason=strChangeReason.split("|")
var strChangeTime="Change Time|变更时间" 
var arrChangeTime=strChangeTime.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strOldStoreCode="Old Code|旧的入库人" 
var arrOldStoreCode=strOldStoreCode.split("|")
var strOldInputQuantity="Old Input Quantity|旧的生产数量" 
var arrOldInputQuantity=strOldInputQuantity.split("|")
var strOldStoreQuantity="Old Store Quantity|旧的入库数量" 
var arrOldStoreQuantity=strOldStoreQuantity.split("|")
var strNewStoreQuantity="New Store Quantity|新的入库数量" 
var arrNewStoreQuantity=strNewStoreQuantity.split("|")
var strOldStoreTime="Old Store Time|旧的入库时间" 
var arrOldStoreTime=strOldStoreTime.split("|")
var strOldStoreType="Old Store Type|旧的入库类型" 
var arrOldStoreType=strOldStoreType.split("|")
var strOldNote="Old Note|注释" 
var arrOldNote=strOldNote.split("|")
var strSubJobs="SubJobs|细分工作单" 
var arrSubJobs=strSubJobs.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeTime.innerText=arrSearchChangeTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeCode.innerText=arrSearchChangeCode[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeType.innerText=arrSearchChangeType[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[0].text=arrSearchChangeTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[1].text=arrSearchChangeEdit[<%=session("language")%>]}catch(e){}
try{document.all.changetype.options[2].text=arrSearchChangeDelete[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_ChangeType.innerText=arrChangeType[<%=session("language")%>]}catch(e){}
try{inner_ChangeCode.innerText=arrChangeCode[<%=session("language")%>]}catch(e){}
try{inner_ChangeReason.innerText=arrChangeReason[<%=session("language")%>]}catch(e){}
try{inner_ChangeTime.innerText=arrChangeTime[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_OldStoreCode.innerText=arrOldStoreCode[<%=session("language")%>]}catch(e){}
try{inner_OldInputQuantity.innerText=arrOldInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_OldStoreQuantity.innerText=arrOldStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_NewStoreQuantity.innerText=arrNewStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_OldStoreTime.innerText=arrOldStoreTime[<%=session("language")%>]}catch(e){}
try{inner_OldStoreType.innerText=arrOldStoreType[<%=session("language")%>]}catch(e){}
try{inner_OldNote.innerText=arrOldNote[<%=session("language")%>]}catch(e){}
try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
}
</script>