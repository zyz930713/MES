<script language="javascript">
var strSearch="Search No Yield|搜索不计良率记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchStoreTime="Store Time|入库时间" 
var arrSearchStoreTime=strSearchStoreTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchCode="Code|入库人" 
var arrSearchCode=strSearchCode.split("|")
var strSearchStoreType="Store Type|入库类型" 
var arrSearchStoreType=strSearchStoreType.split("|")
var strSearchStoreTypeSelect="-- Select Type --|-- 选择类型 --" 
var arrSearchStoreTypeSelect=strSearchStoreTypeSelect.split("|")
var strSearchStoreNormal="Normal|正常工单" 
var arrSearchStoreNormal=strSearchStoreNormal.split("|")
var strSearchStoreRework="Rework|返修工单" 
var arrSearchStoreRework=strSearchStoreRework.split("|")
var strSearchNoYieldStatus="NoYield Status|NoYield状态" 
var arrSearchNoYieldStatus=strSearchNoYieldStatus.split("|")
var strSearchNoYieldStatusSelect="-- Select Status --|-- 选择状态 --" 
var arrSearchNoYieldStatusSelect=strSearchNoYieldStatusSelect.split("|")
var strSearchNoYieldChecked="NoYield Checked|已选择" 
var arrSearchNoYieldChecked=strSearchNoYieldChecked.split("|")
var strSearchNoYieldUnchecked="NoYield Unchecked|未选择" 
var arrSearchNoYieldUnchecked=strSearchNoYieldUnchecked.split("|")

var strBrowse="Browse Store Records|浏览入库记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strNoYield="NoYield|复查" 
var arrNoYield=strNoYield.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strCode="Code|入库人" 
var arrCode=strCode.split("|")
var strInputQuantity="Input Quantity|生产数量" 
var arrInputQuantity=strInputQuantity.split("|")
var strStoreQuantity="Store Quantity|入库数量" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strOntimeYield="OntimeYield|即时良率" 
var arrOntimeYield=strOntimeYield.split("|")
var strStoreTime="Store Time|入库时间" 
var arrStoreTime=strStoreTime.split("|")
var strStoreType="Store Type|入库类型" 
var arrStoreType=strStoreType.split("|")
var strNote="Note|注释" 
var arrNote=strNote.split("|")
var strSubJobs="SubJobs|细分工作单" 
var arrSubJobs=strSubJobs.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchStoreTime.innerText=arrSearchStoreTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_SearchStoreType.innerText=arrSearchStoreType[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[0].text=arrSearchStoreTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[1].text=arrSearchStoreNormal[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[2].text=arrSearchStoreRework[<%=session("language")%>]}catch(e){}
try{inner_SearchNoYieldStatus.innerText=arrSearchNoYieldStatus[<%=session("language")%>]}catch(e){}
try{document.all.NoYieldstatus.options[0].text=arrSearchNoYieldStatusSelect[<%=session("language")%>]}catch(e){}
try{document.all.NoYieldstatus.options[1].text=arrSearchNoYieldChecked[<%=session("language")%>]}catch(e){}
try{document.all.NoYieldstatus.options[2].text=arrSearchNoYieldUnchecked[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_NoYield.innerText=arrNoYield[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_StoreQuantity.innerText=arrStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_OntimeYield.innerText=arrOntimeYield[<%=session("language")%>]}catch(e){}
try{inner_StoreTime.innerText=arrStoreTime[<%=session("language")%>]}catch(e){}
try{inner_StoreType.innerText=arrStoreType[<%=session("language")%>]}catch(e){}
try{inner_Note.innerText=arrNote[<%=session("language")%>]}catch(e){}
try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
}
</script>