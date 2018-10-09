<script language="javascript">
var strSearch="Search Job|查询回收桶工单"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchJobType="Job Type|工单类型" 
var arrSearchJobType=strSearchJobType.split("|")
var strSearchLineName="LineName|线别" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|工厂" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchOptionFactory="Factory|工厂" 
var arrSearchOptionFactory=strSearchOptionFactory.split("|")
var strSearchStatus="Status|状态" 
var arrSearchStatus=strSearchStatus.split("|")
var strSearchOptionStatus="All|所有" 
var arrSearchOptionStatus=strSearchOptionStatus.split("|")
var strSearchCurrentStation="Current Station|当前站" 
var arrSearchCurrentStation=strSearchCurrentStation.split("|")
var strSearchOptionCurrentStation="All|所有" 
var arrSearchOptionCurrentStation=strSearchOptionCurrentStation.split("|")
var strSearchJobStartTime="Job Start Time|工单开始时间" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchCycleTime="Cycle Time|周转时间" 
var arrSearchCycleTime=strSearchCycleTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="to|到" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Job list (Default in past 7 days) in recycle|浏览回收桶内的工单列表（默认7天以内）" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strDelete="Delete Forever|永久删除" 
var arrDelete=strDelete.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|制程" 
var arrPartType=strPartType.split("|")
var strQuantity="Quantity|数量" 
var arrQuantity=strQuantity.split("|")
var strLine="Line|线别" 
var arrLine=strLine.split("|")
var strStartTime="Start Time|开始时间" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Close Time|结束时间" 
var arrCloseTime=strCloseTime.split("|")
var strCycleTime="Cycle Time|周转时间" 
var arrCycleTime=strCycleTime.split("|")
var strStations="Included Stations|站点" 
var arrStations=strStations.split("|")
var strOperators="Stations' Operators|操作工" 
var arrOperators=strOperators.split("|")
var strLineLost="Line Lost|产线损失" 
var arrLineLost=strLineLost.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchJobType.innerText=arrSearchJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchOptionFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchStatus.innerText=arrSearchStatus[<%=session("language")%>]}catch(e){}
try{document.all.status.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{inner_SearchCurrentStation.innerText=arrSearchCurrentStation[<%=session("language")%>]}catch(e){}
try{document.all.currentstation.options[0].text=arrSearchOptionCurrentStation[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchCycleTime.innerText=arrSearchCycleTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Delete.innerText=arrDelete[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Quantity.innerText=arrQuantity[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_Stations.innerText=arrStations[<%=session("language")%>]}catch(e){}
try{inner_Operators.innerText=arrOperators[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>