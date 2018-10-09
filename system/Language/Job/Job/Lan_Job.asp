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
var strSearchJobCloseTime="Close Time|关闭时间" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
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

var strBrowse="Browse Master Job Records|浏览主工单记录" 
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
var strProgress="Progress|入库进度" 
var arrProgress=strProgress.split("|")
var strStock="Stock|入库" 
var arrStock=strStock.split("|")
var strPlaner="Planer|计划人" 
var arrPlaner=strPlaner.split("|")
var strStartQuantity="Start Quantity|计划数量" 
var arrStartQuantity=strStartQuantity.split("|")
var strAssemblyGood="Assembly Good|产线合格品" 
var arrAssemblyGood=strAssemblyGood.split("|")
var strAssemblyYield="Assembly Yield|产线良率" 
var arrAssemblyYield=strAssemblyYield.split("|")
var strLineLost="Line Lost|产线损失" 
var arrLineLost=strLineLost.split("|")
var strFinalGood="Final Good|最终合格品" 
var arrFinalGood=strFinalGood.split("|")
var strConfirmGood="Confirmed Good|实收合格品" 
var arrConfirmGood=strConfirmGood.split("|")
var strFinalYield="Final Yield|最终良率" 
var arrFinalYield=strFinalYield.split("|")
var strFinalScrap="Final Scrap|最终报废" 
var arrFinalScrap=strFinalScrap.split("|")
var strFinalRemain="Remain|剩余" 
var arrFinalRemain=strFinalRemain.split("|")
var strInputTime="Input Time|投产时间" 
var arrInputTime=strInputTime.split("|")
var strLastTime="Last Time|最后时间" 
var arrLastTime=strLastTime.split("|")
var strCycleTime="Cycle Time|周转时间" 
var arrCycleTime=strCycleTime.split("|")
var arrRecords=["No Records","没有记录"]
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchCreateTime.innerText=arrSearchCreateTime[<%=session("language")%>]}catch(e){}
try{inner_SearchJobCloseTime.innerText=arrSearchJobCloseTime[<%=session("language")%>]}catch(e){}
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
try{inner_Progress.innerText=arrProgress[<%=session("language")%>]}catch(e){}
try{inner_Stock.innerText=arrStock[<%=session("language")%>]}catch(e){}
try{inner_Planer.innerText=arrPlaner[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_AssemblyGood.innerText=arrAssemblyGood[<%=session("language")%>]}catch(e){}
try{inner_AssemblyYield.innerText=arrAssemblyYield[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_FinalGood.innerText=arrFinalGood[<%=session("language")%>]}catch(e){}
try{inner_ConfirmGood.innerText=arrConfirmGood[<%=session("language")%>]}catch(e){}
try{inner_FinalYield.innerText=arrFinalYield[<%=session("language")%>]}catch(e){}
try{inner_FinalScrap.innerText=arrFinalScrap[<%=session("language")%>]}catch(e){}
try{inner_FinalRemain.innerText=arrFinalRemain[<%=session("language")%>]}catch(e){}
try{inner_InputTime.innerText=arrInputTime[<%=session("language")%>]}catch(e){}
try{inner_LastTime.innerText=arrLastTime[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>