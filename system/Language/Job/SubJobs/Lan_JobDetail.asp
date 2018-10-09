<script language="javascript">
var strBrowse="Browse a Job|浏览工单" 
var arrBrowse=strBrowse.split("|")
var strSummary="Summary Info|摘要" 
var arrSummary=strSummary.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strPartType="Part Type|制程" 
var arrPartType=strPartType.split("|")
var strStatus="Job Status|工单状态" 
var arrStatus=strStatus.split("|")
var strAction="Job Action|工单动作" 
var arrAction=strAction.split("|")
var strStartTime="Job Start Time|开始时间" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Job Close Time|结束时间" 
var arrCloseTime=strCloseTime.split("|")
var strStartQuantity="Start Quantity|开始数量" 
var arrStartQuantity=strStartQuantity.split("|")
var strGoodQuantity="Good Quantity|成品数量" 
var arrGoodQuantity=strGoodQuantity.split("|")
var strDefectQuantity="Defect Quantity|缺陷数量" 
var arrDefectQuantity=strDefectQuantity.split("|")
var strCycleTime="Cycle Time|周转时间" 
var arrCycleTime=strCycleTime.split("|")
var strDefectHistory="Defectcode Change History|缺陷代码修改历史" 
var arrDefectHistory=strDefectHistory.split("|")
var strStationsInfo="Stations' Info|站点信息" 
var arrStationsInfo=strStationsInfo.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strUpdate="Update|更新" 
var arrUpdate=strUpdate.split("|")
var strStation="Station Name|站点名称" 
var arrStation=strStation.split("|")
var strOperator="Operator|操作工" 
var arrOperator=strOperator.split("|")
var strStationStartTime="Station Start Time|站点开始时间" 
var arrStationStartTime=strStationStartTime.split("|")
var strStationCloseTime="Station Close Time|站点结束时间" 
var arrStationCloseTime=strStationCloseTime.split("|")
var strElapsed="Elapsed|用时" 
var arrElapsed=strElapsed.split("|")
var strMinus="Minus|减" 
var arrMinus=strMinus.split("|")
var strShiftOutTime="Shift Out Time|停线时间" 
var arrShiftOutTime=strShiftOutTime.split("|")
var strShiftInTime="Shift In Time|开线时间" 
var arrShiftInTime=strShiftInTime.split("|")
var strInterval="Interval|间隔" 
var arrInterval=strInterval.split("|")
var strEqual="Equal|等于" 
var arrEqual=strEqual.split("|")
var strCycle="Cycle|周转" 
var arrCycle=strCycle.split("|")
var strCheckAll="Check All|全部选中" 
var arrCheckAll=strCheckAll.split("|")
var strUncheckAll="Uncheck All|全部取消" 
var arrUncheckAll=strUncheckAll.split("|")
var strUpdate="Update|更  新"
var arrUpdate=strUpdate.split("|")
var strReset="Reset|重  设"
var arrReset=strReset.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_Summary.innerText=arrSummary[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_GoodQuantity.innerText=arrGoodQuantity[<%=session("language")%>]}catch(e){}
try{inner_DefectQuantity.innerText=arrDefectQuantity[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_DefectHistory.innerText=arrDefectHistory[<%=session("language")%>]}catch(e){}
try{inner_StationsInfo.innerText=arrStationsInfo[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Update.innerText=arrUpdate[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_StationStartTime.innerText=arrStationStartTime[<%=session("language")%>]}catch(e){}
try{inner_StationCloseTime.innerText=arrStationCloseTime[<%=session("language")%>]}catch(e){}
try{inner_Elapsed.innerText=arrElapsed[<%=session("language")%>]}catch(e){}
try{inner_Minus.innerText=arrMinus[<%=session("language")%>]}catch(e){}
try{inner_ShiftOutTime.innerText=arrShiftOutTime[<%=session("language")%>]}catch(e){}
try{inner_ShiftInTime.innerText=arrShiftInTime[<%=session("language")%>]}catch(e){}
try{inner_Interval.innerText=arrInterval[<%=session("language")%>]}catch(e){}
try{inner_Equal.innerText=arrEqual[<%=session("language")%>]}catch(e){}
try{inner_Cycle.innerText=arrCycle[<%=session("language")%>]}catch(e){}
try{document.all.CheckAll.value=arrCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.UncheckAll.value=arrUncheckAll[<%=session("language")%>]}catch(e){}
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>