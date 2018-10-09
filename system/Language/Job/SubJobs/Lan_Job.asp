<script language="javascript">
var strSearch="Search Job|查询工单"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchJobType="Job Type|工单类型" 
var arrSearchJobType=strSearchJobType.split("|")
var strSearchLineName="Line Name|线别" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|工厂" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchFamily="Family|家族" 
var arrSearchFamily=strSearchFamily.split("|")
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
var strSearchJobStartTime="Start Time|开始时间" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchJobCloseTime="Close Time|关闭时间" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
var strSearchCycleTime="Cycle Time|周转时间" 
var arrSearchCycleTime=strSearchCycleTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Sub Job list (Default in past 7 days)|浏览子工单列表（默认7天以内）" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strDelete="Delete|删除" 
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

var strCurrentStations="Current Stations|当前站点" 
var arrCurrentStations=strCurrentStations.split("|")

var strOperators="Stations' Operators|操作工" 
var arrOperators=strOperators.split("|")
var strLineLost="Line Lost|产线损失" 
var arrLineLost=strLineLost.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")
var arrHoldReport=["Hold Report","工单Hold报告"]
var arrHoldTime=["Hold Time","暂停时间"]
var arrHoldType=["Hold Type","暂停类别"]
var arrHoldReason=["Hold Reason","暂停原因"]
var arr2DCodeReport=["2D Code Report","二维码报表"]
var arr2DCode=["2D Code","二维码"]
var arrLinkTime=["Link Time","绑定时间"]
var arrLinkUser=["Link User","绑定人员"]
var arrSheetNumber=["Sheet Number","分批号"]
var arrReleaseTIme=["Release Time","释放时间"]
var arrPSSTIme=["PSS Time","PSS时间"]
var arrHoldPerson=["Hold Person","暂停人员"]
var arrReleasePerson=["Release Person","释放人员"]
var arrOpCode=["Op Code","操作工"]
var arrHoldStation=["Hold Station","暂停站点"]
var arrHoldJob=["Hold Job","暂停工单"]
var arrReleaseJob=["Release Job","释放工单"]
var arrReleaseReason=["Release Reason","释放原因"]
var arrJobProductRecord=["Job Production Records","工单生产纪录"]
function language()
{
try{inner_JobProductRecord.innerText=arrJobProductRecord[<%=session("language")%>]}catch(e){} 
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchJobType.innerText=arrSearchJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchFamily.innerText=arrSearchFamily[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchOptionFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchStatus.innerText=arrSearchStatus[<%=session("language")%>]}catch(e){}
try{document.all.status.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{document.all.jobtype.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{document.all.jobstatus.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{inner_SearchCurrentStation.innerText=arrSearchCurrentStation[<%=session("language")%>]}catch(e){}
try{document.all.currentstation.options[0].text=arrSearchOptionCurrentStation[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchJobCloseTime.innerText=arrSearchJobCloseTime[<%=session("language")%>]}catch(e){}
try{inner_SearchCycleTime.innerText=arrSearchCycleTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom1.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchTo1.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
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
try{inner_CurrentStations.innerText=arrCurrentStations[<%=session("language")%>]}catch(e){}
try{inner_Operators.innerText=arrOperators[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_HoldReport.innerText=arrHoldReport[<%=session("language")%>]}catch(e){}
try{inner_HoldTime.innerText=arrHoldTime[<%=session("language")%>]}catch(e){}
try{inner_HoldType.innerText=arrHoldType[<%=session("language")%>]}catch(e){}
try{inner_HoldReason.innerText=arrHoldReason[<%=session("language")%>]}catch(e){}
try{inner_2DCodeReport.innerText=arr2DCodeReport[<%=session("language")%>]}catch(e){}
try{inner_2DCode.innerText=arr2DCode[<%=session("language")%>]}catch(e){}
try{inner_linkTime.innerText=arrLinkTime[<%=session("language")%>]}catch(e){}
try{inner_HoldJob.innerText=arrHoldJob[<%=session("language")%>]}catch(e){}
try{td_2DCode.innerText=arr2DCode[<%=session("language")%>]}catch(e){}
try{td_JobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{td_linkUser.innerText=arrLinkUser[<%=session("language")%>]}catch(e){}
try{td_linkTime.innerText=arrLinkTime[<%=session("language")%>]}catch(e){}
try{td_SheetNumber.innerText=arrSheetNumber[<%=session("language")%>]}catch(e){}
try{td_Model.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{td_LineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{td_HoldTime.innerText=arrHoldTime[<%=session("language")%>]}catch(e){}
try{td_ReleaseTime.innerText=arrReleaseTIme[<%=session("language")%>]}catch(e){}
try{td_PSSTime.innerText=arrPSSTIme[<%=session("language")%>]}catch(e){}
try{td_HoldPerson.innerText=arrHoldPerson[<%=session("language")%>]}catch(e){}
try{td_ReleasePerson.innerText=arrReleasePerson[<%=session("language")%>]}catch(e){}
try{td_HoldType.innerText=arrHoldType[<%=session("language")%>]}catch(e){}
try{td_HoldReason.innerText=arrHoldReason[<%=session("language")%>]}catch(e){}
try{td_OpCode.innerText=arrOpCode[<%=session("language")%>]}catch(e){}
try{td_HoldStation.innerText=arrHoldStation[<%=session("language")%>]}catch(e){}

}
</script>